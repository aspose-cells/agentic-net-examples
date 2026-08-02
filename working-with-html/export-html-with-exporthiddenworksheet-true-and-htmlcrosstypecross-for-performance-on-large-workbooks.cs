// Title: Export hidden worksheets to HTML with HtmlCrossType.Cross for high‑performance large workbooks (Aspose.Cells .NET)
// Description: Shows how to save an Aspose.Cells workbook as a single HTML file that includes hidden sheets (ExportHiddenWorksheet = true) and leverages HtmlCrossType.Cross to accelerate export of massive workbooks.
// Keywords: Aspose.Cells HTML export | ExportHiddenWorksheet true | HtmlCrossType.Cross performance | large workbook to HTML | C# Aspose.Cells HtmlSaveOptions | hidden worksheet HTML conversion | .NET Excel to HTML | cross‑cell string optimization
// Common Searches: Aspose.Cells export hidden sheet to HTML | HtmlCrossType.Cross usage example | How to improve HTML export speed for big Excel files | Save entire workbook as one HTML file Aspose.Cells | C# HtmlSaveOptions ExportHiddenWorksheet
// Developer Intent: Generate a single HTML document that contains all worksheets—including hidden ones—and uses the Cross string type to reduce processing time for large Excel files.
// Use Cases: Create a web‑ready audit report that shows data from both visible and hidden worksheets. | Accelerate conversion of workbooks with thousands of rows/columns for online preview. | Produce a single‑file HTML export for archiving or sharing without iterating over each sheet.
// AI Prompts: Provide C# code that exports an Aspose.Cells workbook to HTML, includes hidden worksheets, and sets HtmlCrossType.Cross for performance. | Explain how ExportHiddenWorksheet and HtmlCrossType.Cross affect HTML size and conversion speed in Aspose.Cells. | Show an example of HtmlSaveOptions configuration for large Excel files to minimize memory usage.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to save an Aspose.Cells workbook as a single HTML file that includes hidden sheets (ExportHiddenWorksheet = true) and leverages HtmlCrossType.Cross to accelerate export of massive workbooks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // Add a hidden worksheet with data
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Ensure hidden worksheets are exported
                ExportHiddenWorksheet = true,

                // Use the Cross type for cross‑cell strings to improve performance on large workbooks
                HtmlCrossStringType = HtmlCrossType.Cross,

                // Export the entire workbook (not only the active sheet)
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as HTML with the specified options
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to HTML at: {outputPath}");
        }
    }
}
