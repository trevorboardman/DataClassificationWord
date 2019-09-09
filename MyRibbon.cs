using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;





namespace WordAddIn3
{
    [ComVisible(true)]
    
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        public string ReadDocumentProperty(string propertyName)
        {
            Microsoft.Office.Interop.Word.Document nativeDocument =Globals.ThisAddIn.Application.ActiveDocument;
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

        public void Confidential(Office.IRibbonControl control)
        {

            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Office.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties; 

            if (ReadDocumentProperty("Classification") != null)
            {
                properties["Classification"].Delete();
            }

            properties.Add("Classification", false,
                      Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                      "Confidential");
            
            foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 15;
                footerRange.Text = "Classification: Confidential";
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                headerRange.Font.Size = 15;
                headerRange.Text = "Classification: Confidential";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            
        }
        public void IUO(Office.IRibbonControl control)
        {

            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Office.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties;

            if (ReadDocumentProperty("Classification") != null)
            {
                properties["Classification"].Delete();
            }

            properties.Add("Classification", false,
                      Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                      "Internal");

            foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkBlue;
                footerRange.Font.Size = 15;
                footerRange.Text = "Classification: Internal Use Only";
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkBlue;
                headerRange.Font.Size = 15;
                headerRange.Text = "Classification: Internal Use Only";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        public void RemoveClassificationButton(Office.IRibbonControl control)
        {

            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Office.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties;

            if (ReadDocumentProperty("Classification") != null)
            {
                properties["Classification"].Delete();
            }

            foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Delete();
                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Delete();
            }

        }

        public void PublicUse(Office.IRibbonControl control)
        {

            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Office.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties;

            if (ReadDocumentProperty("Classification") != null)
            {
                properties["Classification"].Delete();
            }

            properties.Add("Classification", false,
                      Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                      "Public");

            foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                footerRange.Font.Size = 15;
                footerRange.Text = "Classification: Public";
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                headerRange.Font.Size = 15;
                headerRange.Text = "Classification: Public";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        public void PrivateUse(Office.IRibbonControl control)
        {

            Microsoft.Office.Interop.Word.Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);
            Office.DocumentProperties properties = (Office.DocumentProperties)vstoDocument.CustomDocumentProperties;

            if (ReadDocumentProperty("Classification") != null)
            {
                properties["Classification"].Delete();
            }

            properties.Add("Classification", false,
                      Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                      "Private");

            foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkYellow;
                footerRange.Font.Size = 15;
                footerRange.Text = "Classification: Private";
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkYellow;
                headerRange.Font.Size = 15;
                headerRange.Text = "Classification: Private";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        public MyRibbon()
        {
           
        }

#region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAddIn3.MyRibbon.xml");
        }

#endregion

#region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

#endregion

#region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

#endregion
    }
}