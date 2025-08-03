//  Copyright 2025 Kasun Nimantha Bamunuarachchi
//  Licensed under the Apache License, Version 2.0 (the "License");
//  you may not use this file except in compliance with the License.
//  You may obtain a copy of the License at
//
//       http://www.apache.org/licenses/LICENSE-2.0
//
//  Unless required by applicable law or agreed to in writing, software
//  distributed under the License is distributed on an "AS IS" BASIS,
//  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//  See the License for the specific language governing permissions and
//  limitations under the License.

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

        /// <summary>
        /// Applies a custom title style to the currently selected text in the Word document.
        /// </summary>
        /// <remarks>This method creates and applies a custom title style to the current selection in the
        /// Word document. The operation is wrapped in an undo record, allowing the user to undo the action as a single
        /// step.</remarks>
        /// <param name="control">The Ribbon control that triggered this action. This parameter is provided by the Office Ribbon framework.</param>
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

        /// <summary>
        /// Applies a custom subtitle style to the currently selected text in the Word document.
        /// </summary>
        /// <remarks>This method starts a custom undo record, applies a predefined subtitle style to the
        /// current selection in the Word document, and then ends the undo record. If no text is selected, the style is
        /// applied to the current cursor position.</remarks>
        /// <param name="control">The Ribbon control that triggered this action.</param>
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

        /// <summary>
        /// Applies a custom paragraph style with placeholder text to the currently selected range in the document.
        /// </summary>
        /// <remarks>This method add/replaces the text in the current selection with predefined placeholder
        /// text  and applies a custom paragraph style to the range. Ensure that a valid selection exists  before
        /// invoking this method.</remarks>
        /// <param name="control">The Ribbon control that triggered this action.</param>
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

        /// <summary>
        /// Applies a custom paragraph style to the currently selected text in the Word document.
        /// </summary>
        /// <remarks>This method starts an undo record to group the style application as a single undoable
        /// action. The custom paragraph style is created and applied to the current selection in the Word
        /// document.</remarks>
        /// <param name="control">The Ribbon control that triggered this action.</param>
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

        /// <summary>
        /// Gets the current <see cref="Word.UndoRecord"/> instance associated with the application.
        /// </summary>
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

        /// <summary>
        /// Creates or retrieves a custom title style named "CustomTitleStyle" in the active Word document.
        /// </summary>
        /// <remarks>If the style "CustomTitleStyle" does not already exist in the active document, this
        /// method creates it with the following attributes: <list type="bullet"> <item><description>Font:
        /// Arial</description></item> <item><description>Font size: 14</description></item> <item><description>Bold:
        /// Enabled</description></item> <item><description>Paragraph alignment: Centered</description></item> </list>
        /// If the style already exists, the existing style is returned.</remarks>
        /// <returns>A <see cref="Word.Style"/> object representing the "CustomTitleStyle" in the active Word document.</returns>
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

        /// <summary>
        /// Creates or retrieves a custom subtitle style in the active Word document.
        /// </summary>
        /// <remarks>This method attempts to create a new style named "CustomSubtitleStyle" in the active
        /// Word document. If the style already exists, it retrieves and returns the existing style instead of creating
        /// a new one. The style is configured with the following attributes: - Font: Arial, size 12, bold. - Paragraph
        /// alignment: Left-aligned.</remarks>
        /// <returns>A <see cref="Word.Style"/> object representing the "CustomSubtitleStyle" in the active Word document.</returns>
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

        /// <summary>
        /// Creates or retrieves a custom paragraph style named "CustomParagraphStyle" in the active Word document.
        /// </summary>
        /// <remarks>If the style "CustomParagraphStyle" does not already exist in the active document,
        /// this method creates it  with predefined formatting: Arial font, size 10, not bold, and justified paragraph
        /// alignment.  If the style already exists, it retrieves the existing style instead of creating a new
        /// one.</remarks>
        /// <returns>A <see cref="Word.Style"/> object representing the "CustomParagraphStyle" in the active Word document.</returns>
        public Word.Style CreateParagraphStyle()
        {
            try
            {
                // Attempt to create a new style
                Word.Style newStyle = Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("CustomParagraphStyle");
                newStyle.Font.Name = "Arial";
                newStyle.Font.Size = 10;
                newStyle.Font.Bold = 0;
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
