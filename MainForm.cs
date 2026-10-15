using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BuildCongestion {
    public class MainForm : Form {

        static readonly Color Red        = Color.FromArgb(0xE3, 0x06, 0x13);
        static readonly Color RedHover   = Color.FromArgb(0xBF, 0x05, 0x10);
        static readonly Color RedSoft    = Color.FromArgb(0xFD, 0xEC, 0xED);
        static readonly Color Bg         = Color.FromArgb(0xF7, 0xF7, 0xF9);
        static readonly Color CardBg     = Color.White;
        static readonly Color LineGrey   = Color.FromArgb(0xE8, 0xE8, 0xEC);
        static readonly Color TextDark   = Color.FromArgb(0x1A, 0x1A, 0x1A);
        static readonly Color TextMuted  = Color.FromArgb(0x8A, 0x8A, 0x90);
        static readonly Color OkGreen    = Color.FromArgb(0x2E, 0x7D, 0x32);
        static readonly Color DisabledFg = Color.FromArgb(0xC0, 0xC0, 0xC6);
        static readonly Color DisabledBg = Color.FromArgb(0xF0, 0xF0, 0xF2);

        TextBox txtMlr, txtPort, txtPla, txtOutput, txtFileName;
        Button  btnMlrBrowse, btnPortBrowse, btnPlaBrowse, btnOutBrowse;
        NumericUpDown numWeek;
        Button  btnGenerate;
        Panel   pbProgress;
        Label   lblStatus;
        int progressPct = 0;
        bool fileNameEditedByUser = false;

        public MainForm(){
            Text = "Huawei Congestion Report  -  v1.0.2";
            ClientSize = new Size(660, 660);
            MinimumSize = new Size(620, 640);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Bg;
            Font = new Font("Segoe UI", 9F);
            DoubleBuffered = true;

            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            BuildUi();
            UpdateChain();
        }

        void BuildUi(){
            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 66;
            header.BackColor = Color.White;
            header.Paint += delegate(object s, PaintEventArgs e){
                Rectangle r = header.ClientRectangle;
                using (SolidBrush b = new SolidBrush(Red))
                    e.Graphics.FillRectangle(b, 0, 0, r.Width, 3);
                using (Pen p = new Pen(LineGrey))
                    e.Graphics.DrawLine(p, 0, r.Height - 1, r.Width, r.Height - 1);
            };

            Label lblTitle = new Label();
            lblTitle.Text = "Huawei Congestion Report";
            lblTitle.ForeColor = TextDark;
            lblTitle.Font = new Font("Segoe UI Semibold", 13.5F, FontStyle.Bold);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 16);
            header.Controls.Add(lblTitle);

            Label lblSub = new Label();
            lblSub.Text = "Port / PLA / MLR  \u2192  one-click Excel workbook   \u00B7   v1.0.2";
            lblSub.ForeColor = TextMuted;
            lblSub.Font = new Font("Segoe UI", 8.25F);
            lblSub.AutoSize = true;
            lblSub.Location = new Point(22, 42);
            header.Controls.Add(lblSub);

            Controls.Add(header);

            Panel card = new Panel();
            card.Location = new Point(18, 84);
            card.Width = ClientSize.Width - 36;
            card.Height = 440;
            card.BackColor = CardBg;
            card.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            card.Paint += delegate(object s, PaintEventArgs e){
                using (Pen p = new Pen(LineGrey))
                    e.Graphics.DrawRectangle(p, 0, 0, card.Width - 1, card.Height - 1);
            };
            Controls.Add(card);

            int y = 14;
            btnMlrBrowse  = AddField(card, "1.  HUAWEI'S MLR FILE", ref y, out txtMlr, false,
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                "Select the Huawei MLR .csv export");
            btnPortBrowse = AddField(card, "2.  PORT DATA FOLDER", ref y, out txtPort, true, null,
                "Folder with Port PM counters");
            btnPlaBrowse  = AddField(card, "3.  PLA DATA FOLDER", ref y, out txtPla, true, null,
                "Folder with PLA PM counters");
            btnOutBrowse  = AddField(card, "4.  OUTPUT FOLDER  (optional)", ref y, out txtOutput, true, null,
                "Leave empty to use the exe folder");

            Label lw = new Label();
            lw.Text = "WEEK NUMBER";
            lw.ForeColor = TextMuted;
            lw.Location = new Point(18, y + 4);
            lw.AutoSize = true;
            lw.Font = new Font("Segoe UI", 7.75F, FontStyle.Bold);
            card.Controls.Add(lw);

            numWeek = new NumericUpDown();
            numWeek.Location = new Point(130, y + 1);
            numWeek.Width = 68;
            numWeek.Minimum = 1;
            numWeek.Maximum = 53;
            numWeek.Value = Core.GetIsoWeek(DateTime.Now);
            numWeek.Font = new Font("Segoe UI", 10F);
            numWeek.BorderStyle = BorderStyle.FixedSingle;
            numWeek.ValueChanged += delegate { UpdateDefaultFileName(); };
            card.Controls.Add(numWeek);

            Label lh = new Label();
            lh.Text = "(default: current ISO week)";
            lh.ForeColor = TextMuted;
            lh.Location = new Point(208, y + 5);
            lh.AutoSize = true;
            lh.Font = new Font("Segoe UI", 8.5F);
            card.Controls.Add(lh);

            y += 44;

            Label lf = new Label();
            lf.Text = "5.  OUTPUT FILE NAME";
            lf.ForeColor = TextMuted;
            lf.Location = new Point(18, y);
            lf.AutoSize = true;
            lf.Font = new Font("Segoe UI", 7.75F, FontStyle.Bold);
            card.Controls.Add(lf);

            txtFileName = new TextBox();
            txtFileName.Location = new Point(18, y + 18);
            txtFileName.Width = card.Width - 36;
            txtFileName.Font = new Font("Segoe UI", 9.5F);
            txtFileName.BorderStyle = BorderStyle.FixedSingle;
            txtFileName.BackColor = Color.White;
            txtFileName.ForeColor = TextDark;
            txtFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtFileName.TextChanged += delegate { fileNameEditedByUser = true; };
            card.Controls.Add(txtFileName);

            y += 52;

            btnGenerate = new Button();
            btnGenerate.Text = "Generate Congestion Report";
            btnGenerate.Location = new Point(18, y);
            btnGenerate.Size = new Size(card.Width - 36, 42);
            btnGenerate.FlatStyle = FlatStyle.Flat;
            btnGenerate.Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold);
            btnGenerate.Cursor = Cursors.Hand;
            btnGenerate.FlatAppearance.BorderSize = 0;
            btnGenerate.FlatAppearance.MouseOverBackColor = RedHover;
            btnGenerate.FlatAppearance.MouseDownBackColor = RedHover;
            btnGenerate.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnGenerate.Click += OnGenerateClick;
            SetButtonEnabled(btnGenerate, false);
            card.Controls.Add(btnGenerate);

            UpdateDefaultFileName();

            pbProgress = new Panel();
            pbProgress.Location = new Point(18, 540);
            pbProgress.Width = ClientSize.Width - 36;
            pbProgress.Height = 10;
            pbProgress.BackColor = LineGrey;
            pbProgress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            pbProgress.Paint += delegate(object s, PaintEventArgs e){
                int w = pbProgress.Width;
                int h = pbProgress.Height;
                int fillW = (int)(w * (double)progressPct / 100.0);
                if (fillW > 0){
                    using (SolidBrush b = new SolidBrush(Red))
                        e.Graphics.FillRectangle(b, 0, 0, fillW, h);
                }
                using (Pen p = new Pen(LineGrey))
                    e.Graphics.DrawRectangle(p, 0, 0, w - 1, h - 1);
            };
            Controls.Add(pbProgress);

            lblStatus = new Label();
            lblStatus.Text = "Ready.  Select the MLR file to begin.";
            lblStatus.ForeColor = TextMuted;
            lblStatus.Location = new Point(18, 556);
            lblStatus.Size = new Size(ClientSize.Width - 36, 42);
            lblStatus.TextAlign = ContentAlignment.MiddleCenter;
            lblStatus.Font = new Font("Segoe UI", 8.75F);
            lblStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(lblStatus);

            Panel footer = new Panel();
            footer.Dock = DockStyle.Bottom;
            footer.Height = 38;
            footer.BackColor = Color.White;
            footer.Paint += delegate(object s, PaintEventArgs e){
                Rectangle r = footer.ClientRectangle;
                using (Pen p = new Pen(LineGrey))
                    e.Graphics.DrawLine(p, 0, 0, r.Width, 0);
                using (SolidBrush b = new SolidBrush(Red))
                    e.Graphics.FillRectangle(b, 0, r.Height - 3, r.Width, 3);
            };

            Label lblCredit = new Label();
            lblCredit.Text = "Developed by Abderrahmane BENHALIMA  \u00B7  BH & Core Transport Planning Engineer  \u00B7  +213 770 900 129  \u00B7  Abderrahmane.BENHALI@DJEZZY.DZ";
            lblCredit.ForeColor = TextMuted;
            lblCredit.Dock = DockStyle.Fill;
            lblCredit.TextAlign = ContentAlignment.MiddleCenter;
            lblCredit.Font = new Font("Segoe UI", 7.5F);
            lblCredit.UseMnemonic = false;
            footer.Controls.Add(lblCredit);
            Controls.Add(footer);

            txtMlr.TextChanged    += delegate { UpdateChain(); };
            txtPort.TextChanged   += delegate { UpdateChain(); };
            txtPla.TextChanged    += delegate { UpdateChain(); };
            txtOutput.TextChanged += delegate { UpdateChain(); };
        }

        void UpdateDefaultFileName(){
            if (fileNameEditedByUser) return;
            if (txtFileName == null) return;
            int w = (int)numWeek.Value;
            string def = Core.DefaultFileName(w);
            bool cur = fileNameEditedByUser;
            fileNameEditedByUser = false;
            txtFileName.Text = def;
            fileNameEditedByUser = cur;
        }

        Button AddField(Panel parent, string label, ref int y, out TextBox tb,
                        bool isFolder, string filter, string placeholder)
        {
            Label lbl = new Label();
            lbl.Text = label;
            lbl.ForeColor = TextMuted;
            lbl.Location = new Point(18, y);
            lbl.AutoSize = true;
            lbl.Font = new Font("Segoe UI", 7.75F, FontStyle.Bold);
            parent.Controls.Add(lbl);

            int btnW = 82;
            int tbW = parent.Width - 36 - btnW - 8;

            tb = new TextBox();
            tb.Location = new Point(18, y + 18);
            tb.Width = tbW;
            tb.Font = new Font("Segoe UI", 9.5F);
            tb.BorderStyle = BorderStyle.FixedSingle;
            tb.BackColor = Color.White;
            tb.ForeColor = TextDark;
            tb.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            parent.Controls.Add(tb);

            Button btn = new Button();
            btn.Text = "Browse";
            btn.Location = new Point(parent.Width - 18 - btnW, y + 17);
            btn.Size = new Size(btnW, 24);
            btn.BackColor = Color.White;
            btn.ForeColor = Red;
            btn.FlatStyle = FlatStyle.Flat;
            btn.Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold);
            btn.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btn.Cursor = Cursors.Hand;
            btn.FlatAppearance.BorderColor = Red;
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = RedSoft;
            btn.FlatAppearance.MouseDownBackColor = RedSoft;
            parent.Controls.Add(btn);

            TextBox capTb = tb;
            bool capF = isFolder;
            string capFilter = filter;
            string capPh = placeholder;
            btn.Click += delegate {
                if (capF){
                    using (FolderBrowserDialog dlg = new FolderBrowserDialog()){
                        dlg.Description = capPh;
                        dlg.ShowNewFolderButton = true;
                        if (dlg.ShowDialog(this) == DialogResult.OK){
                            capTb.Text = dlg.SelectedPath;
                            capTb.SelectionStart = capTb.Text.Length;
                        }
                    }
                } else {
                    using (OpenFileDialog dlg = new OpenFileDialog()){
                        dlg.Filter = capFilter;
                        dlg.Title = capPh;
                        if (dlg.ShowDialog(this) == DialogResult.OK){
                            capTb.Text = dlg.FileName;
                            capTb.SelectionStart = capTb.Text.Length;
                        }
                    }
                }
            };
            y += 52;
            return btn;
        }

        void SetBrowseEnabled(Button btn, bool enabled){
            btn.Enabled = enabled;
            if (enabled){
                btn.ForeColor = Red;
                btn.FlatAppearance.BorderColor = Red;
                btn.FlatAppearance.MouseOverBackColor = RedSoft;
                btn.Cursor = Cursors.Hand;
            } else {
                btn.ForeColor = DisabledFg;
                btn.FlatAppearance.BorderColor = DisabledFg;
                btn.FlatAppearance.MouseOverBackColor = Color.Transparent;
                btn.Cursor = Cursors.Default;
            }
        }

        void SetButtonEnabled(Button btn, bool enabled){
            btn.Enabled = enabled;
            if (enabled){
                btn.BackColor = Red;
                btn.ForeColor = Color.White;
                btn.Cursor = Cursors.Hand;
            } else {
                btn.BackColor = DisabledBg;
                btn.ForeColor = DisabledFg;
                btn.Cursor = Cursors.Default;
            }
        }

        void UpdateChain(){
            bool mlrOk  = txtMlr != null  && File.Exists(txtMlr.Text.Trim());
            bool portOk = txtPort != null && Directory.Exists(txtPort.Text.Trim());
            bool plaOk  = txtPla != null  && Directory.Exists(txtPla.Text.Trim());

            SetBrowseEnabled(btnMlrBrowse, true);
            SetBrowseEnabled(btnPortBrowse, mlrOk);
            txtPort.Enabled = mlrOk || portOk;
            txtPort.BackColor = txtPort.Enabled ? Color.White : DisabledBg;

            SetBrowseEnabled(btnPlaBrowse, mlrOk && portOk);
            txtPla.Enabled = (mlrOk && portOk) || plaOk;
            txtPla.BackColor = txtPla.Enabled ? Color.White : DisabledBg;

            SetBrowseEnabled(btnOutBrowse, true);
            SetButtonEnabled(btnGenerate, mlrOk && portOk && plaOk);
        }

        void SetProgress(int pct){
            if (pbProgress.InvokeRequired){ pbProgress.Invoke(new Action<int>(SetProgress), pct); return; }
            if (pct < 0) pct = 0;
            if (pct > 100) pct = 100;
            progressPct = pct;
            pbProgress.Invalidate();
        }

        void SetStatusText(string s, Color c){
            if (lblStatus.InvokeRequired){ lblStatus.Invoke(new Action<string, Color>(SetStatusText), s, c); return; }
            lblStatus.Text = s;
            lblStatus.ForeColor = c;
        }

        async void OnGenerateClick(object sender, EventArgs e){
            string mlr = txtMlr.Text.Trim();
            string port = txtPort.Text.Trim();
            string pla = txtPla.Text.Trim();
            string outDir = txtOutput.Text.Trim();
            if (string.IsNullOrEmpty(outDir)) outDir = AppDomain.CurrentDomain.BaseDirectory;

            if (!File.Exists(mlr)){ MessageBox.Show(this, "Please select a valid MLR file.", "Missing input", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (!Directory.Exists(port)){ MessageBox.Show(this, "Please select a valid Port data folder.", "Missing input", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (!Directory.Exists(pla)){ MessageBox.Show(this, "Please select a valid PLA data folder.", "Missing input", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            string rawName = txtFileName.Text.Trim();
            if (string.IsNullOrEmpty(rawName)) rawName = Core.DefaultFileName((int)numWeek.Value);
            rawName = Core.SanitizeFileName(rawName);
            if (!rawName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                rawName += ".xlsx";

            int week = (int)numWeek.Value;
            SetButtonEnabled(btnGenerate, false);
            btnGenerate.Text = "Working...";
            SetProgress(0);
            SetStatusText("Starting...", TextMuted);

            string mlrL = mlr, portL = port, plaL = pla, outL = outDir, nameL = rawName;

            Action<int, string> progress = delegate(int pct, string msg){
                SetProgress(pct);
                if (msg != null && msg.Length > 0) SetStatusText(msg, TextMuted);
            };
            Action<string> statusLog = delegate(string s){
                string t = s == null ? "" : s.Trim();
                if (t.Length == 0) return;
                SetStatusText(t, TextMuted);
            };

            try {
                await Task.Run(delegate { Run(mlrL, portL, plaL, outL, nameL, week, progress, statusLog); });
                SetProgress(100);
                SetStatusText("Done.  Report written successfully.", OkGreen);
            } catch (Exception ex){
                SetStatusText("Error: " + ex.Message, Red);
                SetProgress(0);
                MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                btnGenerate.Text = "Generate Congestion Report";
                UpdateChain();
            }
        }

        void Run(string mlrPath, string portDir, string plaDir, string outDir,
                 string fileName, int week,
                 Action<int, string> progress, Action<string> log)
        {
            string[] portCols = {
                "ResourceName","ETH_TX_THROUGHPUT_MAX","ETH_RX_THROUGHPUT_MAX",
                "PORT_TX_BW_UTILIZATION_MAX","PORT_RX_BW_UTILIZATION_MAX",
                "TX_DROP_RATIO","RX_DROP_RATIO"
            };
            string[] plaCols = {
                "ResourceName","ETH_TX_THROUGHPUT_MAX","ETH_RX_THROUGHPUT_MAX",
                "PORT_TX_BW_UTILIZATION_MAX","PORT_RX_BW_UTILIZATION_MAX","TX_DROP_RATIO"
            };

            if (progress != null) progress(1, "Starting...");

            Dictionary<string, LinkCounter> portMap =
                Core.ProcessLinkFolder("Port", portDir, portCols, "RX_DROP_RATIO",
                                       log, progress, 3, 30);

            Dictionary<string, LinkCounter> plaMap =
                Core.ProcessLinkFolder("PLA", plaDir, plaCols, null,
                                       log, progress, 30, 55);

            if (progress != null) progress(55, "Loading MLR file...");
            MlrTable mlr = Core.LoadMlr(mlrPath, log);
            if (mlr == null) throw new Exception("Could not read the MLR file.");
            if (progress != null) progress(58, "MLR loaded: " + mlr.Rows.Count + " rows");

            int cNeSrc  = Core.FindCol(mlr, "Source NE Name", "Source NE", "SourceNE");
            int cNeSnk  = Core.FindCol(mlr, "Sink NE Name", "Sink NE", "SinkNE");
            int cLink   = Core.FindCol(mlr, "Link Name", "LinkName", "Link");
            int cLinkId = Core.FindCol(mlr, "Link ID", "LinkID", "Link Id");
            int cPlaSrc = Core.FindCol(mlr, "Source PLA ID", "Source PLA", "Source PLA-ID", "Source PLAID");
            int cPlaSnk = Core.FindCol(mlr, "Sink PLA ID", "Sink PLA", "Sink PLA-ID", "Sink PLAID");
            int cBrdSrc = Core.FindCol(mlr, "Source Board", "SourceBoard", "Source Board ID");
            int cBrdSnk = Core.FindCol(mlr, "Sink Board", "SinkBoard", "Sink Board ID");
            int cPrtSrc = Core.FindCol(mlr, "Source Port", "SourcePort", "Source Port ID");
            int cPrtSnk = Core.FindCol(mlr, "Sink Port", "SinkPort", "Sink Port ID");

            if (cNeSrc < 0 || cNeSnk < 0 || cLink < 0)
                throw new Exception("MLR file is missing required columns (Source NE / Sink NE / Link Name).");

            if (progress != null) progress(58, "Building S-EPLA index...");
            Dictionary<string, LinkInfo> linkInfoByName =
                new Dictionary<string, LinkInfo>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<LinkInfo>> sEplaByPair =
                new Dictionary<string, List<LinkInfo>>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < mlr.Rows.Count; i++){
                string srcNe = Core.Cell(mlr, i, cNeSrc);
                string snkNe = Core.Cell(mlr, i, cNeSnk);
                string link = Core.Cell(mlr, i, cLink);
                string linkId = cLinkId >= 0 ? Core.Cell(mlr, i, cLinkId) : "";

                string srcSite = Core.NormalizeSite(srcNe);
                string snkSite = Core.NormalizeSite(snkNe);
                string spid = Core.MakeSitePairId(srcSite, snkSite);

                bool isSEpla = false;
                if (cPlaSrc >= 0 && Core.IsSEplaCell(Core.Cell(mlr, i, cPlaSrc))) isSEpla = true;
                if (cPlaSnk >= 0 && Core.IsSEplaCell(Core.Cell(mlr, i, cPlaSnk))) isSEpla = true;

                LinkInfo info = new LinkInfo();
                info.SourceNe = srcNe; info.SinkNe = snkNe;
                info.SourceSite = srcSite; info.SinkSite = snkSite;
                info.SitePairId = spid; info.LinkId = linkId;
                info.LinkName = link; info.IsSEpla = isSEpla;

                if (!string.IsNullOrEmpty(link)){
                    LinkInfo ex;
                    if (!linkInfoByName.TryGetValue(link, out ex) || (isSEpla && !ex.IsSEpla))
                        linkInfoByName[link] = info;
                }
                if (isSEpla && !string.IsNullOrEmpty(spid)){
                    List<LinkInfo> lst;
                    if (!sEplaByPair.TryGetValue(spid, out lst)){
                        lst = new List<LinkInfo>(); sEplaByPair[spid] = lst;
                    }
                    lst.Add(info);
                }
                if ((i & 0x3FF) == 0 && progress != null){
                    int pct = 58 + (12 * i) / Math.Max(1, mlr.Rows.Count);
                    progress(pct, "Building S-EPLA index (" + i + "/" + mlr.Rows.Count + ")");
                }
            }

            Dictionary<string, List<int>> rowsByLink =
                new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < mlr.Rows.Count; i++){
                string rawLink = Core.Cell(mlr, i, cLink);
                string effLink = rawLink;
                if (effLink.Length < Core.LinkNameMinLen)
                    effLink = Core.Cell(mlr, i, cNeSrc) + "-" + Core.Cell(mlr, i, cNeSnk);
                if (string.IsNullOrEmpty(effLink)) continue;
                List<int> lst;
                if (!rowsByLink.TryGetValue(effLink, out lst)){
                    lst = new List<int>(); rowsByLink[effLink] = lst;
                }
                lst.Add(i);
            }

            WilayaIndex wilayas = Core.LoadWilayasFromEmbedded();

            if (progress != null) progress(70, "Matching counters to MLR links...");
            List<CongestionRow> rowsOut = new List<CongestionRow>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, LinkInfo> linkInfoByCong =
                new Dictionary<string, LinkInfo>(StringComparer.OrdinalIgnoreCase);

            int rowCount = mlr.Rows.Count;
            int chunk = Math.Max(1, rowCount / 50);

            for (int i = 0; i < rowCount; i++){
                string rawLink = Core.Cell(mlr, i, cLink);
                string link = rawLink;
                if (link.Length < Core.LinkNameMinLen)
                    link = Core.Cell(mlr, i, cNeSrc) + "-" + Core.Cell(mlr, i, cNeSnk);

                if (!seen.Contains(link)){
                    List<int> group;
                    if (rowsByLink.TryGetValue(link, out group) && group.Count > 0 && i == group[0]){
                        seen.Add(link);

                        LinkCounter bestPla = null, bestPort = null;
                        foreach (int ri in group){
                            string rnSrcNe = Core.Cell(mlr, ri, cNeSrc);
                            string rnSnkNe = Core.Cell(mlr, ri, cNeSnk);
                            string rnPlaSrc = cPlaSrc >= 0 ? Core.Cell(mlr, ri, cPlaSrc) : "";
                            string rnPlaSnk = cPlaSnk >= 0 ? Core.Cell(mlr, ri, cPlaSnk) : "";
                            string rnBrdSrc = cBrdSrc >= 0 ? Core.Cell(mlr, ri, cBrdSrc) : "";
                            string rnBrdSnk = cBrdSnk >= 0 ? Core.Cell(mlr, ri, cBrdSnk) : "";
                            string rnPrtSrc = cPrtSrc >= 0 ? Core.Cell(mlr, ri, cPrtSrc) : "";
                            string rnPrtSnk = cPrtSnk >= 0 ? Core.Cell(mlr, ri, cPrtSnk) : "";

                            string[] plaKeys = {
                                !string.IsNullOrEmpty(rnPlaSrc) ? rnSrcNe + "-" + rnPlaSrc + "(PLA-" + rnPlaSrc + ")" : "",
                                !string.IsNullOrEmpty(rnPlaSnk) ? rnSnkNe + "-" + rnPlaSnk + "(PLA-" + rnPlaSnk + ")" : ""
                            };
                            string[] portKeys = {
                                (!string.IsNullOrEmpty(rnBrdSrc) || !string.IsNullOrEmpty(rnPrtSrc))
                                    ? rnSrcNe + "-" + rnBrdSrc + "-" + rnPrtSrc : "",
                                (!string.IsNullOrEmpty(rnBrdSnk) || !string.IsNullOrEmpty(rnPrtSnk))
                                    ? rnSnkNe + "-" + rnBrdSnk + "-" + rnPrtSnk : ""
                            };

                            LinkCounter pv = PickBest(plaKeys, plaMap);
                            if (pv != null && (bestPla == null || pv.BwPct > bestPla.BwPct)) bestPla = pv;
                            LinkCounter tv = PickBest(portKeys, portMap);
                            if (tv != null && (bestPort == null || tv.BwPct > bestPort.BwPct)) bestPort = tv;
                        }

                        LinkCounter best = bestPla != null ? bestPla : bestPort;

                        CongestionRow crow = new CongestionRow();
                        crow.Wilaya = Core.FindWilaya(link, wilayas);
                        crow.LinkName = link;
                        if (best != null){
                            crow.BwPct = best.BwPct;
                            crow.BwKbps = best.BwKbps;
                            crow.Drop = best.Drop;
                            crow.Range = Core.RangeLabel(best.BwPct);
                        } else {
                            crow.BwPct = Core.NaLabel;
                            crow.BwKbps = Core.NaLabel;
                            crow.Drop = Core.NaLabel;
                            crow.Range = Core.NaLabel;
                        }
                        rowsOut.Add(crow);

                        LinkInfo info;
                        if (!linkInfoByName.TryGetValue(rawLink, out info)){
                            string srcS = Core.NormalizeSite(Core.Cell(mlr, i, cNeSrc));
                            string snkS = Core.NormalizeSite(Core.Cell(mlr, i, cNeSnk));
                            info = new LinkInfo();
                            info.SourceNe = Core.Cell(mlr, i, cNeSrc);
                            info.SinkNe = Core.Cell(mlr, i, cNeSnk);
                            info.SourceSite = srcS; info.SinkSite = snkS;
                            info.SitePairId = Core.MakeSitePairId(srcS, snkS);
                            info.LinkId = ""; info.LinkName = rawLink; info.IsSEpla = false;
                        }
                        linkInfoByCong[link] = info;
                    }
                }

                if (progress != null && (i % chunk == 0)){
                    int pct = 70 + (20 * i) / Math.Max(1, rowCount);
                    progress(pct, "Matching counters (" + i + "/" + rowCount + ")");
                }
            }

            rowsOut.Sort(delegate(CongestionRow a, CongestionRow b){
                double av = (a.BwPct is double) ? (double)a.BwPct : double.NegativeInfinity;
                double bv = (b.BwPct is double) ? (double)b.BwPct : double.NegativeInfinity;
                return bv.CompareTo(av);
            });

            if (progress != null) progress(90, "Filtering NA links related to S-EPLA...");
            List<CongestionRow> kept = new List<CongestionRow>();
            foreach (CongestionRow r in rowsOut){
                bool isNa = r.BwPct is string && (string)r.BwPct == Core.NaLabel;
                if (isNa){
                    LinkInfo info;
                    if (linkInfoByCong.TryGetValue(r.LinkName, out info) && !string.IsNullOrEmpty(info.SitePairId)){
                        List<LinkInfo> lst;
                        if (sEplaByPair.TryGetValue(info.SitePairId, out lst)){
                            bool related = false;
                            foreach (LinkInfo c in lst)
                                if (!string.Equals(c.LinkName, info.LinkName, StringComparison.OrdinalIgnoreCase)){
                                    related = true; break;
                                }
                            if (related){ continue; }
                        }
                    }
                }
                kept.Add(r);
            }
            rowsOut = kept;

            if (progress != null) progress(95, "Writing Excel workbook...");
            string outPath = Path.Combine(outDir, fileName);
            ExcelWriter.Write(outPath, rowsOut, week, log);
            if (progress != null) progress(100, "Excel report saved");
        }

        static LinkCounter PickBest(IEnumerable<string> cands, Dictionary<string, LinkCounter> map){
            LinkCounter best = null;
            foreach (string k in cands){
                if (string.IsNullOrEmpty(k)) continue;
                LinkCounter v;
                if (!map.TryGetValue(k, out v)) continue;
                if (best == null || v.BwPct > best.BwPct) best = v;
            }
            return best;
        }
    }
}