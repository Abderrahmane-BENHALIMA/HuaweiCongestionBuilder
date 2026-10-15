using System;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("Build Congestion - Huawei PM Reporting")]
[assembly: AssemblyProduct("Congestion Huawei")]
[assembly: AssemblyCompany("Abderrahmane BENHALIMA")]
[assembly: AssemblyCopyright("(c) 2026 Abderrahmane BENHALIMA - All rights reserved")]
[assembly: AssemblyVersion("1.0.2.0")]
[assembly: AssemblyFileVersion("1.0.2.0")]

namespace BuildCongestion {
    internal static class Program {
        [STAThread] static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}