using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Word.Application;

namespace CheckLink
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Document doc = app.ActiveDocument;

            if (app.Selection.Hyperlinks.Count == 0)
            {
                // no hyperlink -> no checking
                return;
            }


            // user selected text
            String srcText = app.Selection.Text;
            // hyperlink in the user selected text
            String srcAddress = app.Selection.Hyperlinks[1].Address;

            Debug.WriteLine($"{srcText} : {srcAddress}");

            // Setup searching range
            Range searchRange = doc.Range(app.Selection.End);

            bool hasMatch = true;
            do
            {
                hasMatch = searchRange.Find.Execute(srcText);
                if(hasMatch)
                {
                    // focus on the search result
                    searchRange.Select();

                    // fetch the hyperlink from the search results
                    int cntAddress = searchRange.Hyperlinks.Count;
                    String chkAddress = "";
                    if (cntAddress > 0)
                    {
                        chkAddress = searchRange.Hyperlinks[1].Address;
                    }

                    Debug.WriteLine($"{searchRange.Text} : {chkAddress}");

                    if (string.Compare(chkAddress, srcAddress) != 0)
                    {
                        // Not match - Warn
                        searchRange.HighlightColorIndex = WdColorIndex.wdYellow;
                    }else
                    {
                        // Match - OK
                        searchRange.HighlightColorIndex = WdColorIndex.wdBrightGreen;
                    }

                    // Shrink the searching range
                    searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    Thread.Sleep(250);
                }
                
            } while (hasMatch);
        }
    }
}
