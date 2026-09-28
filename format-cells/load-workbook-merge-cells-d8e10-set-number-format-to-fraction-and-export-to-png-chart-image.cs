// Title: Merge cells D8:E10, apply a fraction number format, and export the worksheet as a PNG image using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing Excel file, merges the range D8:E10, sets the custom number format '# ?/?' on the merged cell, and saves the first worksheet as a PNG image with Aspose.Cells. | Write a C# snippet that uses Aspose.Cells Rendering to apply a fraction format to a merged cell and render the sheet to a PNG file.
// Common Searches: Aspose.Cells C# merge D8:E10 then export worksheet to PNG | How to set '# ?/?' fraction format on merged cells using Aspose.Cells .NET | Render Excel sheet as PNG after merging cells with Aspose.Cells for .NET | C# Aspose.Cells image rendering options for exporting worksheet to PNG
// Tags: cell range merging Aspose.Cells C# | fraction custom number format Aspose.Cells | worksheet PNG rendering Aspose.Cells | image rendering options Aspose.Cells .NET | sheetrender to image C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The program loads input.xlsx, merges cells D8:E10, applies the custom fraction format '# ?/?' to the merged cell, and renders the first worksheet to output.png as a PNG image using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.png";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Merge cells D8:E10 (zero‑based: row 7, column 3, 3 rows, 2 columns)
                worksheet.Cells.Merge(7, 3, 3, 2);

                // Apply a fraction number format to the first cell of the merged range
                var style = worksheet.Cells[7, 3].GetStyle();
                style.Custom = "# ?/?"; // Fraction format like 1 1/2
                worksheet.Cells[7, 3].SetStyle(style);

                // Set image rendering options (default format is PNG)
                var imgOptions = new ImageOrPrintOptions();

                // Render the worksheet to a PNG image
                var sheetRender = new SheetRender(worksheet, imgOptions);
                sheetRender.ToImage(0, outputPath);

                Console.WriteLine($"Worksheet rendered successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
