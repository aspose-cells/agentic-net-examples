using System;
using Aspose.Cells;

namespace AsposeCellsMacroHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink to cell S5 that points to a macro named "MyMacro"
            // The Add method returns the index of the newly created hyperlink
            int hyperlinkIndex = worksheet.Hyperlinks.Add("S5", 1, 1, "MyMacro");

            // Retrieve the hyperlink object to set display text and screen tip
            Hyperlink hyperlink = worksheet.Hyperlinks[hyperlinkIndex];
            hyperlink.TextToDisplay = "Run My Macro";
            hyperlink.ScreenTip = "Click to execute the macro";

            // Enable macros in the workbook (required for macro execution)
            workbook.Settings.EnableMacros = true;

            // Save the workbook
            workbook.Save("MacroHyperlink.xlsx");
        }
    }
}