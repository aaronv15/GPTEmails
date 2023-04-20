using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Windows.Forms;

namespace GPTEmails
{
    public partial class ThisAddIn
    {

        public Ribbon1 Ribbon1;
        private Outlook.Inspectors _inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _inspectors = this.Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            
            if (inspector.CurrentItem is Outlook.MailItem mailItem && mailItem.EntryID == null)
            {
                Ribbon1.newEmail(inspector);
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _inspectors.NewInspector -= Inspectors_NewInspector;
            _inspectors = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
