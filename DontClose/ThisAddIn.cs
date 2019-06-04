using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Windows.Forms;

namespace DontClose
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.ActiveExplorer().WindowState = Outlook.OlWindowState.olMinimized;
        }
        
        private void ThisAddIn_Quit()
        {
            if ((Control.ModifierKeys & Keys.Shift) == Keys.None)
            {
                ProcessStartInfo psOutlook = new ProcessStartInfo("OUTLOOK.EXE", "/recycle");
                psOutlook.WindowStyle = ProcessWindowStyle.Minimized;
                Process.Start(psOutlook);
            }
        }

        #region VSTO 生成的代码
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);
        }

        #endregion
    }
}
