// Title: Export hidden worksheets to HTML with Aspose.Cells for .NET (ExportHiddenWorksheet = true)
// Description: Demonstrates how to create a workbook with a visible and a hidden sheet, configure HtmlSaveOptions (ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false), and save the workbook as a single HTML file that contains the hidden worksheet.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportHiddenWorksheet | hidden worksheet HTML export | save workbook as HTML | include hidden sheets | ExportActiveWorksheetOnly | Aspose.Cells example
// Common Searches: Aspose.Cells export hidden worksheet to HTML | HtmlSaveOptions ExportHiddenWorksheet true example | C# save workbook with hidden sheets as HTML | Include hidden worksheets in HTML output Aspose.Cells | ExportActiveWorksheetOnly vs ExportHiddenWorksheet
// Developer Intent: Include hidden worksheets when converting a workbook to HTML.
// Use Cases: Generate a web‑ready report that shows data from both visible and hidden tabs for audit trails. | Provide an online preview of a spreadsheet where supplemental hidden sheets must be visible to end users. | Create documentation that preserves all worksheet content, including hidden sections, in HTML format.
// AI Prompts: Show a C# code snippet that exports a workbook to HTML with hidden worksheets using Aspose.Cells. | Explain the interaction between ExportHiddenWorksheet and ExportActiveWorksheetOnly in HtmlSaveOptions. | Give step‑by‑step instructions to save an Aspose.Cells workbook as HTML while ensuring hidden sheets are included.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with a visible and a hidden sheet, configure HtmlSaveOptions (ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false), and save the workbook as a single HTML file that contains the hidden worksheet.
    public class ExportHiddenWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // First worksheet (visible)
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Visible Data");

                // Add a second worksheet and hide it
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Hidden Data");
                hiddenSheet.IsVisible = false; // Mark worksheet as hidden

                // Configure HTML save options to include hidden worksheets
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true,      // Ensure hidden sheets are exported
                    ExportActiveWorksheetOnly = false // Export the whole workbook
                };

                // Determine output path
                string outputPath = "output_with_hidden.html";

                // Save the workbook as HTML; hidden sheet will be included in the output
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHiddenWorksheetDemo.Run();
        }
    }
}
