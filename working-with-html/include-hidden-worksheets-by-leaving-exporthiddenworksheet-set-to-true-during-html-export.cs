// Title: Export Hidden Worksheets to HTML with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions (ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false), ensure the output folder exists, and save the workbook as an HTML file that includes hidden worksheets using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportHiddenWorksheet | hidden worksheet export | Excel to HTML conversion | ExportActiveWorksheetOnly | save workbook as HTML | include hidden sheets
// Common Searches: Aspose.Cells export hidden worksheets to HTML C# | HtmlSaveOptions ExportHiddenWorksheet example | How to include hidden sheets when saving Excel as HTML | Export entire workbook with hidden sheets using Aspose.Cells | C# save workbook to HTML with hidden worksheets
// Developer Intent: Generate an HTML file from an Excel workbook that retains both visible and hidden worksheets.
// Use Cases: Create an audit‑ready HTML report that shows data from hidden sheets. | Provide a web preview of an Excel file where supplemental hidden information must be visible. | Automate documentation pipelines that convert Excel workbooks to HTML while preserving all worksheets.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML and include hidden worksheets. | Explain the impact of ExportHiddenWorksheet and ExportActiveWorksheetOnly on the HTML output. | Give a step‑by‑step guide to save a workbook with hidden sheets to HTML and create the output directory if it does not exist.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions (ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false), ensure the output folder exists, and save the workbook as an HTML file that includes hidden worksheets using Aspose.Cells for .NET.
    public class ExportHiddenWorksheetsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook with a default visible sheet
                Workbook workbook = new Workbook();
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Visible Data");

                // Add a hidden worksheet and put some data in it
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Hidden Data");
                hiddenSheet.IsVisible = false; // hide the sheet

                // Set HTML save options to export hidden worksheets
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true,          // include hidden sheets
                    ExportActiveWorksheetOnly = false      // export the whole workbook
                };

                // Ensure the output directory exists
                string outputPath = "output_with_hidden.html";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML; hidden worksheet will be included
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHiddenWorksheetsDemo.Run();
        }
    }
}
