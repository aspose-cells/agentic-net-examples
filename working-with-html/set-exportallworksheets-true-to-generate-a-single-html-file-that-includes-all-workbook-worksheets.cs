// Title: Export every worksheet of an Excel workbook to one HTML file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export multiple worksheets to a single HTML page | How to save an entire Excel workbook as one HTML document using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example in C# | Generate HTML from a workbook with gridlines, headings and Base64 images in C# | Export workbook to HTML with evaluated formulas using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false | C# export entire workbook to single HTML file | HTML export with gridlines and row/column headings Aspose.Cells | embed images as Base64 in Aspose.Cells HTML output | evaluate formulas before HTML conversion Aspose.Cells

using System;
using Aspose.Cells;

namespace ExportAllWorksheetsToHtml
{
    // The example creates a workbook with two worksheets, fills them with data and formulas, calculates the formulas, and then uses HtmlSaveOptions (ExportActiveWorksheetOnly = false, ExportGridLines = true, ExportRowColumnHeadings = true, ExportImagesAsBase64 = true) to save the entire workbook as a single HTML file that includes all worksheets, gridlines, headings, and embedded images.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook with two worksheets
                Workbook workbook = new Workbook();

                // Worksheet 1 - default name "Sheet1"
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Products";
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Price");
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["B2"].PutValue(1.20);
                sheet1.Cells["A3"].PutValue("Banana");
                sheet1.Cells["B3"].PutValue(0.80);

                // Worksheet 2 - add a new sheet
                int sheet2Index = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Summary";
                sheet2.Cells["A1"].PutValue("Total Products");
                sheet2.Cells["B1"].Formula = "=COUNTA(Products!A2:A100)";
                sheet2.Cells["A2"].PutValue("Average Price");
                sheet2.Cells["B2"].Formula = "=AVERAGE(Products!B2:B100)";

                // Calculate formulas so that the HTML will contain evaluated values
                workbook.CalculateFormula();

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export all worksheets (default behavior)
                    ExportActiveWorksheetOnly = false,

                    // Keep gridlines and headings for better readability
                    ExportGridLines = true,
                    ExportRowColumnHeadings = true,

                    // Embed images as Base64 strings
                    ExportImagesAsBase64 = true
                };

                // Define output path
                string outputPath = "AllWorksheetsExport.html";

                // Save the workbook as a single HTML file containing all worksheets
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully exported to HTML file: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during the export process:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
