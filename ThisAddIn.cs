using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Reflection;


namespace WordAddIn3
{
    public partial class ThisAddIn
    {

        
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            
             this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
       private void DocumentBeforeSave()
        {
           

        }

        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Microsoft.Office.Core.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties;
            string messageBoxText = "We noticed you haven't classified your document, would you like to opt out of classifying this document?";
            string caption = "Classify Your Document";

            object oBasic = Application.WordBasic;

            object fIsAutoSave =
             oBasic.GetType().InvokeMember(
                "IsAutosaveEvent",
             BindingFlags.GetProperty,
                 null, oBasic, null);

            MessageBoxButtons button = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Warning;

            if (ReadDocumentProperty2("Classification") == null && !(int.Parse(fIsAutoSave.ToString()) == 1))
            {
                DialogResult result = MessageBox.Show(messageBoxText, caption, button, icon);
                switch (result)
                {
                    case DialogResult.Yes:
                        properties.Add("Classification", false,
                       Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                       "Not Classified");
                        break;
                    case DialogResult.No:
                        MessageBox.Show("Please add a classification and try saving the document again", "Classify", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Cancel = true;
                        break;

                }

            }

        }

        public string ReadDocumentProperty2(string propertyName)
        {
            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Office.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
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


    public class DocumentClass 
    {


    }
}
