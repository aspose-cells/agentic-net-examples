// Title: Export an Excel workbook to HTML while preserving merged cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export .xlsx to .html preserving merged cells | How to keep colspan and rowspan when converting Excel to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions example for merged cell support in .NET | Save Excel workbook as HTML with merged ranges using Aspose.Cells C# | Convert Excel file to HTML with merged cells retained using Aspose.Cells library
// Tags: Aspose.Cells HTML export merged cells | C# HtmlSaveOptions SaveFormat.Html | preserve colspan rowspan Aspose.Cells | convert .xlsx to .html using Aspose.Cells | export Excel to HTML with merged ranges C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // Demonstrates loading an existing .xlsx file with Aspose.Cells, configuring HtmlSaveOptions (merged cells are retained by default), and saving the workbook as an .html file, including error handling for missing files and runtime exceptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel workbook
                var workbook = new Workbook(inputPath);

                // Configure HTML save options (merged cells are exported by default)
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
