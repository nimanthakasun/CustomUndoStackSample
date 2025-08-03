using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace CustomUndoStackSample
{
    [ComVisible(true)]
    public class CustomUndoStackRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private string placeholderText = "This text was added by the Ribbon of CustomUndoSample. " +
            "If you click the dropdown arrow next to the Undo button, you'll notice that no " +
            "custom undo record has been created for this specific operation. " +
            "Instead, each internal step of the process appears individually in the Undo stack - " +
            "even though the entire action was triggered by a single click. \n" +
            "This behavior can be confusing and cluttered from a user experience" +
            " perspective, as it exposes implementation details that users shouldn't have to deal with. " +
            "Ideally, all those internal actions should be grouped into a single, " +
            "clean undo entry to improve usability and keep the Undo history meaningful.\n" +
            "Try Uncommenting \"MyUndoRecord.StartCustomRecord(\"My Para Style With Placeholder\");\" line and " +
            "\"MyUndoRecord.EndCustomRecord();\" lines in the \"OnParagraphButton()\" method.";

        public CustomUndoStackRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CustomUndoStackSample.CustomUndoStackRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnTitleButton(Office.IRibbonControl control)
        {
            MyUndoRecord.StartCustomRecord("My Title Style");
            try
            {
                Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
                currentRange.set_Style(CreateTitleStyle());
            }
            finally
            {
                MyUndoRecord.EndCustomRecord();
            }
        }

        public void OnSubTitleButton(Office.IRibbonControl control)
        {
            MyUndoRecord.StartCustomRecord("My Subtitle Style");
            try
            {
                Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
                currentRange.set_Style(CreateSubTitleStyle());
            }
            finally
            {
                MyUndoRecord.EndCustomRecord();
            }
            
        }

        public void OnParagraphButton(Office.IRibbonControl control)
        {
            //MyUndoRecord.StartCustomRecord("My Para Style With Placeholder");
            try
            {
                Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
                currentRange.Text = placeholderText;
                currentRange.set_Style(CreateParagraphStyle());
            }
            finally
            {
                //MyUndoRecord.EndCustomRecord();
            }

        }

        public void OnStyleButton(Office.IRibbonControl control)
        {
            MyUndoRecord.StartCustomRecord("My Paragraph Style");
            try
            {
                Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
                currentRange.set_Style(CreateParagraphStyle());
            }
            finally
            {
                MyUndoRecord.EndCustomRecord();
            }

        }
        #endregion

        #region Helpers

        Word.UndoRecord MyUndoRecord
        {
            get
            {
                return Globals.ThisAddIn.Application.UndoRecord;
            }
        }
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

        public Word.Style CreateTitleStyle()
        {
            try
            {
                // Attempt to create a new style
                Word.Style newStyle = Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("CustomTitleStyle");
                newStyle.Font.Name = "Arial";
                newStyle.Font.Size = 14;
                newStyle.Font.Bold = 1;
                newStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                return newStyle;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // If the style already exists, we can just return it
                return Globals.ThisAddIn.Application.ActiveDocument.Styles["CustomTitleStyle"];
            }
        }

        public Word.Style CreateSubTitleStyle()
        {
            try
            {
                // Attempt to create a new style
                Word.Style newStyle = Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("CustomSubtitleStyle");
                newStyle.Font.Name = "Arial";
                newStyle.Font.Size = 12;
                newStyle.Font.Bold = 1;
                newStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                return newStyle;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // If the style already exists, return it
                return Globals.ThisAddIn.Application.ActiveDocument.Styles["CustomSubtitleStyle"];
            }
        }

        public Word.Style CreateParagraphStyle()
        {
            try
            {
                // Attempt to create a new style
                Word.Style newStyle = Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("CustomParagraphStyle");
                newStyle.Font.Name = "Arial";
                newStyle.Font.Size = 10;
                newStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                return newStyle;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // If the style already exists, return it
                return Globals.ThisAddIn.Application.ActiveDocument.Styles["CustomParagraphStyle"];
            }
        }

        #endregion
    }
}
