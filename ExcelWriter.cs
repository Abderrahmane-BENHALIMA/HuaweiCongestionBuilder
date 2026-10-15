using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace BuildCongestion {
    public static class ExcelWriter {
        public static string Write(string outPath, List<CongestionRow> rows, int weekNum, Action<string> log){
            Type xlT = Type.GetTypeFromProgID("Excel.Application");
            if (xlT == null) throw new InvalidOperationException("Microsoft Excel is not installed.");
            object appObj = null, wbObj = null;
            try {
                appObj = Activator.CreateInstance(xlT);
                dynamic app = appObj;
                app.Visible = false; app.DisplayAlerts = false; app.ScreenUpdating = false;
                wbObj = app.Workbooks.Add();
                dynamic wb = wbObj;

                dynamic wsSum = wb.Worksheets.Item(1); wsSum.Name = "Summary";
                dynamic ws70  = wb.Worksheets.Add(Type.Missing, wsSum); ws70.Name = ">70%";
                dynamic wsAll = wb.Worksheets.Add(Type.Missing, ws70);  wsAll.Name = "ALL_W" + weekNum;
                dynamic wsThr = wb.Worksheets.Add(Type.Missing, wsAll); wsThr.Name = "Throughput Statistics";

                string[] cols = { "Wilaya","Link Name","BW Utilization (%)","BW Utilization (kbps)","Drop (%)","Range (%)" };
                int nData = rows.Count;

                for (int c=0;c<cols.Length;c++) wsAll.Cells[1,c+1].Value2 = cols[c];
                dynamic hAll = wsAll.Range(wsAll.Cells[1,1], wsAll.Cells[1,6]);
                hAll.Font.Bold=true; hAll.Interior.Color=Rgb(0x37,0x56,0x23); hAll.Font.Color=Rgb(0xFF,0xFF,0xFF);

                for (int i=0;i<nData;i++){
                    CongestionRow r = rows[i]; int xl = i+2;
                    wsAll.Cells[xl,1].Value2 = r.Wilaya == null ? "" : r.Wilaya;
                    wsAll.Cells[xl,2].Value2 = r.LinkName == null ? "" : r.LinkName;
                    wsAll.Cells[xl,3].Value2 = (r.BwPct  is double) ? r.BwPct  : (object)Core.NaLabel;
                    wsAll.Cells[xl,4].Value2 = (r.BwKbps is double) ? r.BwKbps : (object)Core.NaLabel;
                    wsAll.Cells[xl,5].Value2 = (r.Drop   is double) ? r.Drop   : (object)Core.NaLabel;
                    wsAll.Cells[xl,6].Value2 = r.Range == null ? "" : r.Range;
                }
                wsAll.Cells[1,8].Value2 = "_rank_gt70";
                wsAll.Cells[1,9].Value2 = "_rank_thr";
                wsAll.Cells[1,10].Value2 = "_rank_drop";
                for (int i=0;i<nData;i++){
                    int xl = i+2;
                    string fH = (xl==2)
                        ? "=IF(OR($F2=\"[70%-80%]\",$F2=\"[80%-90%]\",$F2=\"[90%-100%]\"),1,0)"
                        : "=IF(OR($F"+xl+"=\"[70%-80%]\",$F"+xl+"=\"[80%-90%]\",$F"+xl+"=\"[90%-100%]\"),N($H"+(xl-1)+")+1,N($H"+(xl-1)+"))";
                    string fI = (xl==2) ? "=IF(N($D2)>"+Core.ThrKbpsThreshold+",1,0)"
                        : "=IF(N($D"+xl+")>"+Core.ThrKbpsThreshold+",N($I"+(xl-1)+")+1,N($I"+(xl-1)+"))";
                    string fJ = (xl==2) ? "=IF(N($E2)>"+Core.DropThreshold+",1,0)"
                        : "=IF(N($E"+xl+")>"+Core.DropThreshold+",N($J"+(xl-1)+")+1,N($J"+(xl-1)+"))";
                    wsAll.Cells[xl,8].Formula = fH;
                    wsAll.Cells[xl,9].Formula = fI;
                    wsAll.Cells[xl,10].Formula = fJ;
                }
                wsAll.Columns["H:J"].Hidden = true;
                wsAll.Columns[1].ColumnWidth = 22; wsAll.Columns[2].ColumnWidth = 48;
                wsAll.Columns[3].ColumnWidth = 18; wsAll.Columns[4].ColumnWidth = 22;
                wsAll.Columns[5].ColumnWidth = 12; wsAll.Columns[6].ColumnWidth = 14;

                List<string> wilayas = new List<string>();
                foreach (CongestionRow r in rows){
                    string w = string.IsNullOrEmpty(r.Wilaya) ? Core.NotRecordedLabel : r.Wilaya;
                    if (!wilayas.Contains(w)) wilayas.Add(w);
                }
                var cNA = new Dictionary<string,int>(); var c70 = new Dictionary<string,int>();
                var c00 = new Dictionary<string,int>(); var c80 = new Dictionary<string,int>();
                var c90 = new Dictionary<string,int>();
                foreach (string w in wilayas){ cNA[w]=0; c70[w]=0; c00[w]=0; c80[w]=0; c90[w]=0; }
                foreach (CongestionRow r in rows){
                    string w = string.IsNullOrEmpty(r.Wilaya) ? Core.NotRecordedLabel : r.Wilaya;
                    string rng = r.Range == null ? "" : r.Range;
                    if (rng == Core.NaLabel) cNA[w]++;
                    else if (rng == "[00%-70%]") c00[w]++;
                    else if (rng == "[70%-80%]") c70[w]++;
                    else if (rng == "[80%-90%]") c80[w]++;
                    else if (rng == "[90%-100%]") c90[w]++;
                }
                wilayas.Sort(delegate(string a, string b){
                    return (c70[b]+c80[b]+c90[b]).CompareTo(c70[a]+c80[a]+c90[a]);
                });

                wsSum.Cells[1,1].Value2 = "Range BW Utilization % W" + weekNum;
                wsSum.Range(wsSum.Cells[1,1], wsSum.Cells[1,5]).Merge();
                wsSum.Cells[1,1].Font.Bold = true; wsSum.Cells[1,1].Font.Size = 12;
                wsSum.Cells[1,1].Font.Color = Rgb(0xFF,0xFF,0xFF); wsSum.Cells[1,1].Interior.Color = Rgb(0x1F,0x4E,0x78);
                wsSum.Range(wsSum.Cells[1,1], wsSum.Cells[1,5]).HorizontalAlignment = -4108;

                string[] sumH = { "Wilaya","[70%-80%]","[80%-90%]","[90%-100%]","Total" };
                for (int c=0;c<sumH.Length;c++){
                    wsSum.Cells[2,c+1].Value2 = sumH[c];
                    wsSum.Cells[2,c+1].HorizontalAlignment = -4108;
                }
                dynamic hSum = wsSum.Range(wsSum.Cells[2,1], wsSum.Cells[2,5]);
                hSum.Font.Bold=true; hSum.Interior.Color=Rgb(0x1F,0x4E,0x78); hSum.Font.Color=Rgb(0xFF,0xFF,0xFF);

                for (int i=0;i<wilayas.Count;i++){
                    string w = wilayas[i]; int xl = i+3;
                    wsSum.Cells[xl,1].Value2 = w;
                    wsSum.Cells[xl,2].Value2 = c70[w];
                    wsSum.Cells[xl,3].Value2 = c80[w];
                    wsSum.Cells[xl,4].Value2 = c90[w];
                    wsSum.Cells[xl,5].Formula = "=SUM(B"+xl+":D"+xl+")";
                }
                int totRow = wilayas.Count + 3;
                wsSum.Cells[totRow,1].Value2 = "Total";
                for (int c=2;c<=5;c++){
                    char cl = (char)('A' + c - 1);
                    wsSum.Cells[totRow,c].Formula = "=SUM("+cl+"3:"+cl+(totRow-1)+")";
                }
                wsSum.Range(wsSum.Cells[totRow,1], wsSum.Cells[totRow,5]).Font.Bold = true;
                wsSum.Range(wsSum.Cells[totRow,1], wsSum.Cells[totRow,5]).Interior.Color = Rgb(0xD9,0xE1,0xF2);
                int ratioRow = totRow + 1;
                wsSum.Cells[ratioRow,1].Value2 = "Congestion Ratio %"; wsSum.Cells[ratioRow,1].Font.Bold = true;
                for (int c=2;c<=5;c++){
                    char cl = (char)('A' + c - 1);
                    wsSum.Cells[ratioRow,c].Formula =
                        "=IFERROR("+cl+totRow+"/(COUNTA('ALL_W"+weekNum+"'!$B:$B)-COUNTIF('ALL_W"+weekNum+"'!$F:$F,\"NA\")-1)*100,0)";
                    wsSum.Cells[ratioRow,c].NumberFormat = "0.000000000";
                    wsSum.Cells[ratioRow,c].Font.Bold = true;
                }
                wsSum.Columns[1].ColumnWidth = 22;
                for (int c=2;c<=4;c++) wsSum.Columns[c].ColumnWidth = 14;
                wsSum.Columns[5].ColumnWidth = 16;

                int chartCol = 24;
                string[] chH = { "Wilaya","Non recorded","[00%-70%]","[70%-80%]","[80%-90%]","[90%-100%]" };
                for (int c=0;c<chH.Length;c++) wsSum.Cells[1,chartCol+c].Value2 = chH[c];
                List<string> chartWil = new List<string>(wilayas);
                chartWil.Sort(delegate(string a,string b){
                    int ta = cNA[a]+c00[a]+c70[a]+c80[a]+c90[a];
                    int tb = cNA[b]+c00[b]+c70[b]+c80[b]+c90[b];
                    return tb.CompareTo(ta);
                });
                for (int i=0;i<chartWil.Count;i++){
                    string w = chartWil[i]; int xl = i+2;
                    wsSum.Cells[xl,chartCol].Value2 = w;
                    wsSum.Cells[xl,chartCol+1].Value2 = cNA[w];
                    wsSum.Cells[xl,chartCol+2].Value2 = c00[w];
                    wsSum.Cells[xl,chartCol+3].Value2 = c70[w];
                    wsSum.Cells[xl,chartCol+4].Value2 = c80[w];
                    wsSum.Cells[xl,chartCol+5].Value2 = c90[w];
                }
                wsSum.Range(wsSum.Cells[1,chartCol], wsSum.Cells[1,chartCol+5]).EntireColumn.Hidden = true;

                int n = chartWil.Count;
                dynamic chObj = wsSum.ChartObjects().Add(640, 10, 900, Math.Max(600, n*20+180));
                dynamic chart = chObj.Chart;
                try { chart.PlotVisibleOnly = false; } catch { }

                chart.ChartType = 58;
                try { chart.Axes(1).ReversePlotOrder = true; } catch { }
                try { chart.Axes(1).TickLabels.Orientation = -4105; } catch { }

                dynamic seriesColl = chart.SeriesCollection();
                while (((int)seriesColl.Count) > 0) seriesColl.Item(1).Delete();
                string[] sN = { "Non recorded","[00%-70%]","[70%-80%]","[80%-90%]","[90%-100%]" };
                string[] sC = { "#808080","#32DA0C","#F7F703","#F59F09","#F90713" };
                for (int i=0;i<sN.Length;i++){
                    dynamic s = seriesColl.NewSeries();
                    s.Name = sN[i];
                    s.Values  = wsSum.Range(wsSum.Cells[2,chartCol+1+i], wsSum.Cells[n+1,chartCol+1+i]);
                    s.XValues = wsSum.Range(wsSum.Cells[2,chartCol],     wsSum.Cells[n+1,chartCol]);
                    try { s.Format.Fill.ForeColor.RGB = ColorFromHex(sC[i]); } catch { }
                    s.HasDataLabels = true; s.DataLabels().Font.Size = 8;
                }
                chart.HasTitle = true; chart.ChartTitle.Text = "Range BW Utilization % W" + weekNum;
                chart.HasLegend = true; chart.Legend.Position = -4107;

                ws70.Cells[1,1].Value2 = "Links > 70% - W" + weekNum;
                ws70.Range(ws70.Cells[1,1], ws70.Cells[1,6]).Merge();
                ws70.Cells[1,1].Font.Bold = true; ws70.Cells[1,1].Font.Size = 12;
                ws70.Cells[1,1].Font.Color = Rgb(0xFF,0xFF,0xFF); ws70.Cells[1,1].Interior.Color = Rgb(0x1F,0x4E,0x78);
                ws70.Cells[1,1].HorizontalAlignment = -4108;
                ws70.Cells[1,1].VerticalAlignment   = -4108;
                for (int c=0;c<cols.Length;c++){
                    ws70.Cells[2,c+1].Value2 = cols[c];
                    ws70.Cells[2,c+1].Font.Bold = true;
                    ws70.Cells[2,c+1].Interior.Color = Rgb(0x1F,0x4E,0x78);
                    ws70.Cells[2,c+1].Font.Color = Rgb(0xFF,0xFF,0xFF);
                    ws70.Cells[2,c+1].HorizontalAlignment = -4108;
                }
                List<CongestionRow> r70 = new List<CongestionRow>();
                foreach (CongestionRow r in rows)
                    if (r.Range == "[70%-80%]" || r.Range == "[80%-90%]" || r.Range == "[90%-100%]") r70.Add(r);
                for (int k=0;k<r70.Count;k++){
                    int xl = k+3;
                    for (int c=0;c<6;c++){
                        char cl = (char)('A'+c);
                        ws70.Cells[xl,c+1].Formula =
                            "=IFERROR(INDEX('ALL_W"+weekNum+"'!$"+cl+":$"+cl+",MATCH("+(k+1)+",'ALL_W"+weekNum+"'!$H:$H,0)),\"\")";
                    }
                }
                ws70.Columns[1].ColumnWidth = 22; ws70.Columns[2].ColumnWidth = 48;
                ws70.Columns[3].ColumnWidth = 18; ws70.Columns[4].ColumnWidth = 22;
                ws70.Columns[5].ColumnWidth = 12; ws70.Columns[6].ColumnWidth = 14;

                wsThr.Cells[1,1].Value2 = "Max Link Throughput (>2Gbps)";
                wsThr.Range(wsThr.Cells[1,1], wsThr.Cells[1,6]).Merge();
                wsThr.Cells[1,1].Font.Bold = true; wsThr.Cells[1,1].Font.Color = Rgb(0xFF,0xFF,0xFF);
                wsThr.Cells[1,1].Interior.Color = Rgb(0x20,0x38,0x64);
                wsThr.Cells[1,1].HorizontalAlignment = -4108;
                wsThr.Cells[1,1].VerticalAlignment   = -4108;
                for (int c=0;c<cols.Length;c++){
                    wsThr.Cells[2,c+1].Value2 = cols[c];
                    wsThr.Cells[2,c+1].Font.Bold = true;
                    wsThr.Cells[2,c+1].Interior.Color = Rgb(0x20,0x38,0x64);
                    wsThr.Cells[2,c+1].Font.Color = Rgb(0xFF,0xFF,0xFF);
                    wsThr.Cells[2,c+1].HorizontalAlignment = -4108;
                }
                List<CongestionRow> rThr = new List<CongestionRow>();
                foreach (CongestionRow r in rows) if (r.BwKbps is double && (double)r.BwKbps > Core.ThrKbpsThreshold) rThr.Add(r);
                rThr.Sort(delegate(CongestionRow a, CongestionRow b){ return ((double)b.BwKbps).CompareTo((double)a.BwKbps); });
                for (int k=0;k<rThr.Count;k++){
                    int xl = k+3;
                    for (int c=0;c<6;c++){
                        char cl = (char)('A'+c);
                        wsThr.Cells[xl,c+1].Formula =
                            "=IFERROR(INDEX('ALL_W"+weekNum+"'!$"+cl+":$"+cl+",MATCH("+(k+1)+",'ALL_W"+weekNum+"'!$I:$I,0)),\"\")";
                    }
                }
                int gap = 8;
                wsThr.Cells[1,gap].Value2 = "Max Link with Drop (>2%)";
                wsThr.Range(wsThr.Cells[1,gap], wsThr.Cells[1,gap+5]).Merge();
                wsThr.Cells[1,gap].Font.Bold = true; wsThr.Cells[1,gap].Font.Color = Rgb(0xFF,0xFF,0xFF);
                wsThr.Cells[1,gap].Interior.Color = Rgb(0x84,0x3C,0x0C);
                wsThr.Cells[1,gap].HorizontalAlignment = -4108;
                wsThr.Cells[1,gap].VerticalAlignment   = -4108;
                for (int c=0;c<cols.Length;c++){
                    wsThr.Cells[2,gap+c].Value2 = cols[c];
                    wsThr.Cells[2,gap+c].Font.Bold = true;
                    wsThr.Cells[2,gap+c].Interior.Color = Rgb(0x84,0x3C,0x0C);
                    wsThr.Cells[2,gap+c].Font.Color = Rgb(0xFF,0xFF,0xFF);
                    wsThr.Cells[2,gap+c].HorizontalAlignment = -4108;
                }
                List<CongestionRow> rDrop = new List<CongestionRow>();
                foreach (CongestionRow r in rows) if (r.Drop is double && (double)r.Drop > Core.DropThreshold) rDrop.Add(r);
                rDrop.Sort(delegate(CongestionRow a, CongestionRow b){ return ((double)b.Drop).CompareTo((double)a.Drop); });
                for (int k=0;k<rDrop.Count;k++){
                    int xl = k+3;
                    for (int c=0;c<6;c++){
                        char cl = (char)('A'+c);
                        wsThr.Cells[xl,gap+c].Formula =
                            "=IFERROR(INDEX('ALL_W"+weekNum+"'!$"+cl+":$"+cl+",MATCH("+(k+1)+",'ALL_W"+weekNum+"'!$J:$J,0)),\"\")";
                    }
                }
                int[] widths = { 22,48,18,22,12,14 };
                for (int c=0;c<6;c++){ wsThr.Columns[c+1].ColumnWidth = widths[c]; wsThr.Columns[gap+c].ColumnWidth = widths[c]; }

                wsAll.Activate();
                if (File.Exists(outPath)) File.Delete(outPath);
                wb.SaveAs(outPath, 51);
                wb.Close(false);
                app.Quit();
                if (log != null) log("Excel report: " + Path.GetFileName(outPath));
                return outPath;
            } finally {
                try { if (wbObj != null) Marshal.ReleaseComObject(wbObj); } catch {}
                try { if (appObj != null) Marshal.ReleaseComObject(appObj); } catch {}
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }
        static int Rgb(int r, int g, int b){ return r | (g << 8) | (b << 16); }
        static int ColorFromHex(string hex){
            hex = hex.TrimStart('#');
            int r = Convert.ToInt32(hex.Substring(0,2),16);
            int g = Convert.ToInt32(hex.Substring(2,2),16);
            int b = Convert.ToInt32(hex.Substring(4,2),16);
            return Rgb(r,g,b);
        }
    }
}