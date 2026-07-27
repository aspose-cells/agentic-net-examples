// Title: Include Hidden Worksheets in HTML Export with Aspose.Cells for .NET (ExportHiddenWorksheet = true)
// Description: Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions (ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false), and save the workbook as HTML so that hidden worksheets are rendered in the output.
// Keywords: Aspose.Cells HTML export hidden worksheets | ExportHiddenWorksheet true | ExportActiveWorksheetOnly false | C# Aspose.Cells HTML conversion | include hidden sheets in HTML output
// Common Searches: Aspose.Cells export hidden sheet to HTML | HtmlSaveOptions ExportHiddenWorksheet example C# | Save entire workbook as HTML including hidden worksheets | How to export hidden worksheets with Aspose.Cells .NET
// Developer Intent: Generate an HTML file from an Excel workbook that contains both visible and hidden worksheets using Aspose.Cells for .NET.
// Use Cases: Web preview of a workbook where hidden sheets hold supplemental data that must be visible in the HTML report. | Automated conversion pipeline that preserves all worksheets, regardless of visibility, when exporting Excel to HTML. | Creating downloadable HTML versions of Excel files for documentation or archival purposes while keeping hidden content accessible.
// AI Prompts: Show C# code that exports an Excel workbook to HTML with hidden worksheets included using Aspose.Cells. | Explain the impact of ExportHiddenWorksheet and ExportActiveWorksheetOnly settings when converting Excel to HTML. | Provide a step‑by‑step guide to configure HtmlSaveOptions for full‑workbook HTML export in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions (ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false), and save the workbook as HTML so that hidden worksheets are rendered in the output.
    public class ExportHiddenWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ----- Visible worksheet -----
                Worksheet visibleSheet = workbook.Worksheets[0]; // default first sheet
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Data from visible sheet");

                // ----- Hidden worksheet -----
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Data from hidden sheet");
                hiddenSheet.IsVisible = false; // mark as hidden

                // Configure HTML save options to include hidden worksheets
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true,          // ensure hidden sheets are exported
                    ExportActiveWorksheetOnly = false      // export the whole workbook
                };

                // Save the workbook to HTML; hidden worksheet will be part of the output
                workbook.Save("ExportHiddenWorksheet_Enabled.html", saveOptions);
                Console.WriteLine("Workbook exported successfully to ExportHiddenWorksheet_Enabled.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHiddenWorksheetDemo.Run();
        }
    }
}
