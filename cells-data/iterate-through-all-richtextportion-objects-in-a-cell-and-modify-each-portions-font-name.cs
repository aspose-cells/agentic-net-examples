// Title: Set the font of cell A1 to Arial in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file, change the font of cell A1 to Arial with Aspose.Cells in C#, and save the workbook to a new location. | Create a custom style that specifies a font name and apply it to a single cell in a worksheet using the Aspose.Cells .NET API. | Verify that a cell contains text, then programmatically set its font to Arial and write the updated workbook to disk with Aspose.Cells C#.
// Common Searches: asp.net how to change font of a specific Excel cell using Aspose.Cells | c# set cell A1 font to Arial in existing workbook Aspose.Cells | apply custom style to a single cell and save workbook with Aspose.Cells .NET | verify input file exists before modifying Excel cell font Aspose.Cells | save modified Excel file to different folder using Aspose.Cells C#
// Tags: set cell font Aspose.Cells C# | apply style to single Excel cell Aspose.Cells | load and save workbook Aspose.Cells .NET | check cell text before formatting Aspose.Cells | create custom font style Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRichTextExample
{
    // The example checks that the source .xlsx file exists, loads it with Aspose.Cells, accesses cell A1 on the first worksheet, creates a style with the Arial font, applies the style to the cell if it contains text, ensures the output directory is present, and saves the modified workbook to the specified path while handling possible exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = @"C:\Path\To\Input.xlsx";
            string outputPath = @"C:\Path\To\Output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Get the cell that may contain rich text (e.g., A1)
                var cell = worksheet.Cells["A1"];

                // If the cell contains any text, apply the desired font to the whole text
                if (!string.IsNullOrEmpty(cell.StringValue))
                {
                    // Create a style with the desired font properties
                    var style = workbook.CreateStyle();
                    style.Font.Name = "Arial";

                    // Apply the style to the entire cell text
                    cell.SetStyle(style);
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
