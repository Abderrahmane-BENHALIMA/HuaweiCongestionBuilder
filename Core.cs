using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace BuildCongestion {
    public class LinkCounter { public double BwPct; public double BwKbps; public double Drop; }
    public class LinkInfo {
        public string SourceNe=""; public string SinkNe="";
        public string SourceSite=""; public string SinkSite="";
        public string SitePairId=""; public string LinkId="";
        public string LinkName=""; public bool IsSEpla;
    }
    public class CongestionRow {
        public string Wilaya=""; public string LinkName="";
        public object BwPct; public object BwKbps; public object Drop;
        public string Range="";
    }
    public class CsvHint {
        public Encoding Enc; public char Sep; public int Skip;
        public CsvHint(Encoding e, char s, int k){Enc=e;Sep=s;Skip=k;}
        public bool SameAs(CsvHint o){ return o!=null && Enc.WebName==o.Enc.WebName && Sep==o.Sep && Skip==o.Skip; }
    }
    public class MlrTable {
        public string[] Headers = new string[0];
        public List<string[]> Rows = new List<string[]>();
    }
    public class WilayaIndex {
        public Dictionary<char,List<KeyValuePair<string,string>>> Site = new Dictionary<char,List<KeyValuePair<string,string>>>();
        public Dictionary<char,List<KeyValuePair<string,string>>> Code = new Dictionary<char,List<KeyValuePair<string,string>>>();
    }

    public static class Core {
        public const int LinkNameMinLen = 7;
        public static readonly string[] Ranges = { "[70%-80%]","[80%-90%]","[90%-100%]" };
        public static readonly string[] AllRanges = { "[00%-70%]","[70%-80%]","[80%-90%]","[90%-100%]" };
        public const string NaLabel = "NA";
        public const string NotRecordedLabel = "(not recorded)";
        public const double ThrKbpsThreshold = 2000000;
        public const double DropThreshold = 2;

        static readonly Regex SiteRe   = new Regex(@"^(O\d+[A-Z]\d+)", RegexOptions.IgnoreCase);
        static readonly Regex SEplaRe  = new Regex(@"S[\-_]?EPLA", RegexOptions.IgnoreCase);
        static readonly Regex NormRe   = new Regex(@"[^a-z0-9_]");

        public static string Normalize(string s){
            if (s == null) return "";
            return NormRe.Replace(s.Trim().ToLowerInvariant(), "");
        }
        public static string NormalizeSite(string ne){
            if (string.IsNullOrEmpty(ne)) return "";
            Match m = SiteRe.Match(ne.Trim());
            return m.Success ? m.Groups[1].Value.ToUpperInvariant() : "";
        }
        public static string MakeSitePairId(string a, string b){
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return "";
            return (string.CompareOrdinal(a,b) <= 0) ? (a+"|"+b) : (b+"|"+a);
        }
        public static int GetIsoWeek(DateTime d){
            DayOfWeek dow = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(d);
            if (dow >= DayOfWeek.Monday && dow <= DayOfWeek.Wednesday) d = d.AddDays(3);
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }
        public static string DefaultFileName(int weekNum){
            return "Congestion Huawei Week " + weekNum + " (" + DateTime.Now.ToString("yyyy-MM-dd") + ")";
        }
        public static string SanitizeFileName(string name){
            if (string.IsNullOrEmpty(name)) return "";
            char[] bad = Path.GetInvalidFileNameChars();
            foreach (char c in bad) name = name.Replace(c, '_');
            return name.Trim();
        }
        public static string DetectSep(string line){
            if (line.IndexOf('\t') >= 0) return "\t";
            if (line.IndexOf(';')  >= 0) return ";";
            return ",";
        }
        public static string[] SplitCsvLine(string line, char sep){
            List<string> res = new List<string>();
            bool inQ = false; StringBuilder sb = new StringBuilder();
            for (int i=0;i<line.Length;i++){
                char c = line[i];
                if (c == '"') {
                    if (inQ && i+1 < line.Length && line[i+1]=='"'){ sb.Append('"'); i++; }
                    else inQ = !inQ;
                } else if (c == sep && !inQ){ res.Add(sb.ToString()); sb.Length=0; }
                else sb.Append(c);
            }
            res.Add(sb.ToString()); return res.ToArray();
        }
        public static Encoding[] Encodings(){
            return new Encoding[]{
                new UTF8Encoding(true), new UTF8Encoding(false),
                Encoding.Unicode, Encoding.BigEndianUnicode,
                Encoding.GetEncoding("iso-8859-1")
            };
        }
        public static readonly char[] Separators = { '\t',';',',' };
        public static readonly int[] SkipRowsOptions = { 1,0,2,3,4,5 };

        public static double? ToNumOrNull(string s){
            double v; if (string.IsNullOrEmpty(s)) return null;
            s = s.Trim().Replace("%","").Replace(" ","");
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;
            return null;
        }

        static List<Dictionary<string,string>> TryRead(string path, CsvHint t, string reqNorm, out bool ok){
            ok = false;
            try {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (StreamReader sr = new StreamReader(fs, t.Enc)) {
                    List<string> lines = new List<string>(); string line; int c = 0;
                    while ((line = sr.ReadLine()) != null){ if (c++ < t.Skip) continue; lines.Add(line); }
                    if (lines.Count < 1) return null;
                    string[] hc = SplitCsvLine(lines[0], t.Sep);
                    string[] hcols = new string[hc.Length];
                    for (int i=0;i<hc.Length;i++) hcols[i] = Normalize(hc[i]);
                    bool found = false;
                    for (int i=0;i<hcols.Length;i++) if (hcols[i]==reqNorm){ found=true; break; }
                    if (!found) return null;
                    List<Dictionary<string,string>> rows = new List<Dictionary<string,string>>();
                    for (int i=1;i<lines.Count;i++){
                        if (string.IsNullOrEmpty(lines[i]) || lines[i].Trim().Length==0) continue;
                        string[] cells = SplitCsvLine(lines[i], t.Sep);
                        Dictionary<string,string> d = new Dictionary<string,string>();
                        for (int k=0;k<hcols.Length && k<cells.Length;k++) if (!d.ContainsKey(hcols[k])) d[hcols[k]] = cells[k];
                        rows.Add(d);
                    }
                    ok = true; return rows;
                }
            } catch { return null; }
        }
        public static List<Dictionary<string,string>> ReadHuaweiCsv(string path, string requiredCol, out string status, ref CsvHint hint){
            status = ""; string reqNorm = Normalize(requiredCol);
            List<CsvHint> tries = new List<CsvHint>();
            if (hint != null) tries.Add(hint);
            foreach (Encoding e in Encodings()) foreach (char s in Separators) foreach (int k in SkipRowsOptions){
                CsvHint cand = new CsvHint(e,s,k);
                if (hint != null && cand.SameAs(hint)) continue;
                tries.Add(cand);
            }
            foreach (CsvHint t in tries){
                bool ok; var rows = TryRead(path, t, reqNorm, out ok);
                if (!ok) continue;
                string sd = (t.Sep=='\t') ? "\\t" : t.Sep.ToString();
                status = "enc=" + t.Enc.WebName + ", sep='" + sd + "', skiprows=" + t.Skip + ", rows=" + rows.Count;
                hint = t; return rows;
            }
            status = "no (encoding, separator, header-line) combination matched";
            return null;
        }

        public static Dictionary<string,LinkCounter> ProcessLinkFolder(
            string label, string folder, string[] inputCols, string rxDropCol,
            Action<string> log,
            Action<int, string> progress,
            int pctStart, int pctEnd)
        {
            var empty = new Dictionary<string,LinkCounter>();
            if (!Directory.Exists(folder)){
                if (log != null) log(label + ": folder not found.");
                if (progress != null) progress(pctEnd, label + ": folder not found");
                return empty;
            }
            string[] allowed = new string[]{ "", ".csv", ".txt", ".tsv", ".dat" };
            List<string> files = new List<string>();
            string[] all = Directory.GetFiles(folder);
            Array.Sort(all, StringComparer.OrdinalIgnoreCase);
            foreach (string f in all){
                string ext = Path.GetExtension(f).ToLowerInvariant();
                for (int i=0;i<allowed.Length;i++) if (allowed[i]==ext){ files.Add(f); break; }
            }
            if (files.Count == 0){
                if (log != null) log(label + ": no data files.");
                if (progress != null) progress(pctEnd, label + ": no data files");
                return empty;
            }
            if (log != null) log(label + ": " + files.Count + " file(s) to process");
            if (progress != null) progress(pctStart, "Reading " + label + " data (0/" + files.Count + ")");

            string rn = Normalize("ResourceName");
            string txBw = Normalize("PORT_TX_BW_UTILIZATION_MAX");
            string rxBw = Normalize("PORT_RX_BW_UTILIZATION_MAX");
            string txThr = Normalize("ETH_TX_THROUGHPUT_MAX");
            string rxThr = Normalize("ETH_RX_THROUGHPUT_MAX");
            string txDrop = Normalize("TX_DROP_RATIO");
            string rxDrop = (rxDropCol != null) ? Normalize(rxDropCol) : null;

            var map = new Dictionary<string,LinkCounter>();
            CsvHint hint = null; int idx = 0;
            foreach (string path in files){
                idx++; string status;
                var rows = ReadHuaweiCsv(path, "ResourceName", out status, ref hint);
                if (rows != null){
                    foreach (var r in rows){
                        string rnVal;
                        if (!r.TryGetValue(rn, out rnVal)) continue;
                        if (string.IsNullOrEmpty(rnVal) || rnVal=="nan" || rnVal=="None") continue;
                        string s1,s2,s3,s4,s5,s6;
                        double? txBwV  = r.TryGetValue(txBw,   out s1) ? ToNumOrNull(s1) : null;
                        double? rxBwV  = r.TryGetValue(rxBw,   out s2) ? ToNumOrNull(s2) : null;
                        double? txThrV = r.TryGetValue(txThr,  out s3) ? ToNumOrNull(s3) : null;
                        double? rxThrV = r.TryGetValue(rxThr,  out s4) ? ToNumOrNull(s4) : null;
                        double? txDrV  = r.TryGetValue(txDrop, out s5) ? ToNumOrNull(s5) : null;
                        double? rxDrV  = (rxDrop != null && r.TryGetValue(rxDrop, out s6)) ? ToNumOrNull(s6) : null;
                        if (txBwV == null && rxBwV == null) continue;
                        double bwPct  = Math.Max(txBwV.HasValue ? txBwV.Value : 0, rxBwV.HasValue ? rxBwV.Value : 0);
                        double bwKbps = Math.Max(txThrV.HasValue ? txThrV.Value : 0, rxThrV.HasValue ? rxThrV.Value : 0);
                        double drop = (rxDrop == null)
                            ? (txDrV.HasValue ? txDrV.Value : 0)
                            : Math.Max(txDrV.HasValue ? txDrV.Value : 0, rxDrV.HasValue ? rxDrV.Value : 0);
                        rnVal = rnVal.Trim();
                        LinkCounter ex;
                        if (!map.TryGetValue(rnVal, out ex) || bwPct > ex.BwPct){
                            LinkCounter lc = new LinkCounter(); lc.BwPct=bwPct; lc.BwKbps=bwKbps; lc.Drop=drop;
                            map[rnVal] = lc;
                        }
                    }
                }
                if (log != null) log(label + " " + idx + "/" + files.Count + "  (" + Path.GetFileName(path) + ")");
                if (progress != null){
                    int pct = pctStart + (pctEnd - pctStart) * idx / files.Count;
                    progress(pct, "Reading " + label + " data (" + idx + "/" + files.Count + ")");
                }
            }
            if (log != null) log(label + " done: " + map.Count + " unique ResourceName(s)");
            return map;
        }

        public static MlrTable LoadMlr(string path, Action<string> log){
            if (!File.Exists(path)){ if (log != null) log("MLR not found"); return null; }
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sr = new StreamReader(fs, Encoding.UTF8)){
                string header = null;
                for (int i=0;i<11;i++) sr.ReadLine();
                header = sr.ReadLine();
                if (header == null){ if (log != null) log("MLR has no header row"); return null; }
                char sep = DetectSep(header)[0];
                string[] hc = SplitCsvLine(header, sep);
                string[] headers = new string[hc.Length];
                for (int i=0;i<hc.Length;i++) headers[i] = hc[i].Trim();
                MlrTable t = new MlrTable(); t.Headers = headers;
                string line;
                while ((line = sr.ReadLine()) != null){
                    if (string.IsNullOrEmpty(line) || line.Trim().Length==0) continue;
                    t.Rows.Add(SplitCsvLine(line, sep));
                }
                if (log != null) log("MLR loaded: " + t.Rows.Count + " rows");
                return t;
            }
        }
        public static int FindCol(MlrTable t, params string[] cands){
            foreach (string c in cands) for (int i=0;i<t.Headers.Length;i++)
                if (string.Equals(t.Headers[i], c, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }
        public static string Cell(MlrTable t, int row, int col){
            if (col < 0 || row < 0 || row >= t.Rows.Count) return "";
            var r = t.Rows[row]; if (col >= r.Length) return "";
            return r[col] == null ? "" : r[col].Trim();
        }
        public static bool IsSEplaCell(string v){ return v != null && SEplaRe.IsMatch(v); }

        static void AddBucket(Dictionary<char,List<KeyValuePair<string,string>>> b, string p, string w){
            if (string.IsNullOrEmpty(p)) return;
            char k = p[0];
            if (!b.ContainsKey(k)) b[k] = new List<KeyValuePair<string,string>>();
            b[k].Add(new KeyValuePair<string,string>(p,w));
        }
        public static WilayaIndex LoadWilayasFromEmbedded(){
            WilayaIndex idx = new WilayaIndex();
            string[] lines = WilayaData.Csv.Replace("\r\n","\n").Split('\n');
            if (lines.Length < 2) return idx;
            char sep = DetectSep(lines[0])[0];
            string[] header = SplitCsvLine(lines[0], sep);
            int iCode=-1,iSite=-1,iWil=-1;
            for (int i=0;i<header.Length;i++){
                string h = header[i].Trim().ToLowerInvariant();
                if (h=="code") iCode=i; else if (h=="site code") iSite=i; else if (h=="wilaya") iWil=i;
            }
            for (int i=1;i<lines.Length;i++){
                if (string.IsNullOrEmpty(lines[i]) || lines[i].Trim().Length==0) continue;
                string[] c = SplitCsvLine(lines[i], sep);
                string w  = (iWil >=0 && iWil <c.Length) ? c[iWil].Trim()  : "";
                string sc = (iSite>=0 && iSite<c.Length) ? c[iSite].Trim() : "";
                string cc = (iCode>=0 && iCode<c.Length) ? c[iCode].Trim() : "";
                if (string.IsNullOrEmpty(w)) continue;
                AddBucket(idx.Site, sc, w); AddBucket(idx.Code, cc, w);
            }
            foreach (char k in new List<char>(idx.Site.Keys))
                idx.Site[k].Sort(delegate(KeyValuePair<string,string> a, KeyValuePair<string,string> b){ return b.Key.Length.CompareTo(a.Key.Length); });
            foreach (char k in new List<char>(idx.Code.Keys))
                idx.Code[k].Sort(delegate(KeyValuePair<string,string> a, KeyValuePair<string,string> b){ return b.Key.Length.CompareTo(a.Key.Length); });
            return idx;
        }
        public static string FindWilaya(string link, WilayaIndex idx){
            if (string.IsNullOrEmpty(link)) return "";
            char first = link[0];
            List<KeyValuePair<string,string>> sb;
            if (idx.Site.TryGetValue(first, out sb))
                foreach (var kv in sb) if (link.StartsWith(kv.Key, StringComparison.OrdinalIgnoreCase)) return kv.Value;
            List<KeyValuePair<string,string>> cb;
            if (idx.Code.TryGetValue(first, out cb))
                foreach (var kv in cb) if (link.StartsWith(kv.Key, StringComparison.OrdinalIgnoreCase)) return kv.Value;
            foreach (var bucket in idx.Site.Values) foreach (var kv in bucket) if (link.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0) return kv.Value;
            foreach (var bucket in idx.Code.Values) foreach (var kv in bucket) if (link.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0) return kv.Value;
            return "";
        }
        public static string RangeLabel(object v){
            if (v == null) return "";
            if (!(v is double)) return "";
            double d = (double)v; if (double.IsNaN(d)) return "";
            if (d < 70) return "[00%-70%]";
            if (d < 80) return "[70%-80%]";
            if (d < 90) return "[80%-90%]";
            return "[90%-100%]";
        }
    }
}