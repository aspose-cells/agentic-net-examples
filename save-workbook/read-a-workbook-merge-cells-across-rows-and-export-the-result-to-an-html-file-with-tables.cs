// Title: Merge cells vertically in an Excel worksheet and export the result to an HTML file with tables using Aspose.Cells for .NET (C#)
// Common Searches: c# merge a vertical range of cells and save as html with Aspose.Cells | asp.net export only the active worksheet to html using HtmlSaveOptions | how to embed images as base64 when converting Excel to html with Aspose.Cells | aspose.cells convert merged cells to html table in c# | save excel workbook as html file with merged cells using Aspose.Cells
// Tags: vertical cell merging Aspose.Cells C# | export worksheet to html Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly C# | embed images base64 Aspose.Cells HTML | save workbook as html table Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndExport
{
    // The example loads an Excel workbook, merges cells A1:A3 vertically on the first worksheet, sets a value for the merged cell, and then saves only that active worksheet as an HTML file. HtmlSaveOptions are configured to export the worksheet as an HTML table and embed any images as Base64 strings.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path for the output HTML file
            string outputPath = "output.html";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Example: Merge cells vertically across rows (A1:A3)
            // This creates a single cell that spans rows 1 to 3 in column A
            sheet.Cells.Merge(0, 0, 3, 1); // (firstRow, firstColumn, totalRows, totalColumns)

            // Optionally, you can set a value for the merged cell
            sheet.Cells[0, 0].PutValue("Merged Cells");

            // Save the workbook as an HTML file with tables
            // The HtmlSaveOptions can be customized if needed
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.ExportActiveWorksheetOnly = true; // Export only the active sheet
            saveOptions.ExportImagesAsBase64 = true;      // Embed images as Base64

            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("Workbook has been processed and saved as HTML.");
        }
    }
}
