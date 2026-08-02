// Title: Convert Excel to HTML with Aspose.Cells (C#) – Include Hidden Worksheets
// Description: Load an .xlsx file, configure HtmlSaveOptions to export hidden sheets, and save the workbook as an HTML document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells HTML export C# | Export hidden worksheets | HtmlSaveOptions ExportHiddenWorksheet | Excel to HTML conversion | C# workbook to HTML
// Common Searches: Aspose.Cells export hidden sheets to HTML | C# convert Excel file to HTML with hidden worksheets | default HtmlSaveOptions example Aspose.Cells | how to include hidden worksheets in HTML output
// Developer Intent: Generate an HTML file from an Excel workbook while preserving any hidden worksheets.
// Use Cases: Publish a web‑ready view of a workbook that contains internal analysis sheets. | Create HTML reports from templates where hidden tabs hold supplemental data. | Automate batch conversion of multiple Excel files to HTML without losing hidden content.
// AI Prompts: Show C# code that converts an Excel workbook to HTML and keeps hidden worksheets using Aspose.Cells. | Explain the default behavior of HtmlSaveOptions.ExportHiddenWorksheet and when to set it. | Write a script that processes a folder of .xlsx files, converting each to HTML while exporting hidden sheets.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Load an .xlsx file, configure HtmlSaveOptions to export hidden sheets, and save the workbook as an HTML document using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options with default settings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Ensure hidden worksheets are exported (default is true, set explicitly for clarity)
            saveOptions.ExportHiddenWorksheet = true;

            // Save the workbook as an HTML file
            string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook has been converted to HTML and saved to: {outputPath}");
        }
    }
}
