// Title: Export Hidden Worksheets to HTML with Aspose.Cells for .NET (ExportHiddenWorksheet = true)
// Description: Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions to include hidden worksheets, and save the entire workbook as an HTML file so that hidden data appears in the output.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportHiddenWorksheet | hidden worksheet HTML export | ExportActiveWorksheetOnly | save workbook as HTML | include hidden sheets | Aspose.Cells .NET | HTML conversion
// Common Searches: Aspose.Cells export hidden sheet to HTML | HtmlSaveOptions ExportHiddenWorksheet example C# | How to include hidden worksheets when saving as HTML | ExportActiveWorksheetOnly false hidden sheets | HTML output missing hidden worksheet Aspose.Cells
// Developer Intent: Generate an HTML representation of a workbook that contains data from both visible and hidden worksheets.
// Use Cases: Publish a complete web report that shows supplemental data stored in hidden sheets. | Archive a workbook in HTML format while preserving every sheet for compliance review. | Email an HTML snapshot of a workbook where hidden sheets hold critical information that must be visible to recipients.
// AI Prompts: Show how to export only selected hidden worksheets to HTML using Aspose.Cells. | Provide a C# snippet that sets ExportHiddenWorksheet true and customizes the output folder and file name. | Explain the interaction between ExportHiddenWorksheet and ExportActiveWorksheetOnly when saving to HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHiddenWorksheetHtmlDemo
{
    // Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions to include hidden worksheets, and save the entire workbook as an HTML file so that hidden data appears in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // First worksheet – visible
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Second worksheet – hidden
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to export hidden worksheets
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true,   // Ensure hidden sheets are included
                ExportActiveWorksheetOnly = false // Export the whole workbook
            };

            // Save the workbook to HTML; hidden sheet content will appear in the output
            workbook.Save("Workbook_WithHiddenSheet.html", htmlOptions);

            Console.WriteLine("HTML file generated with hidden worksheet content included.");
        }
    }
}
