// Title: Remove All Shapes from an Excel Workbook and Export to CSV with Aspose.Cells for .NET
// Description: Loads a workbook, removes every drawing object (pictures, charts, text boxes) from each worksheet using `RemoveAllDrawingObjects`, then saves the first worksheet as a CSV file. Includes file‑existence validation and exception handling for reliable batch processing.
// Keywords: Aspose.Cells remove shapes | delete drawing objects .NET | export worksheet to CSV | clean CSV output Aspose.Cells | RemoveAllDrawingObjects example
// Common Searches: how to delete all shapes in Excel with Aspose.Cells | Aspose.Cells .NET remove drawing objects before CSV export | export first sheet to CSV after stripping images | batch convert Excel to CSV without charts
// Developer Intent: Strip every shape from all worksheets and save the first sheet as a CSV file.
// Use Cases: Prepare data extracts from templates that contain placeholder images before importing into a database. | Batch‑convert legacy Excel reports to CSV while eliminating embedded graphics to reduce file size. | Generate clean CSV files for downstream analytics when source workbooks include charts or logos.
// AI Prompts: Create a C# method using Aspose.Cells that removes all drawing objects from each worksheet and saves a chosen sheet as CSV. | Provide a .NET code snippet that checks the input path, calls RemoveAllDrawingObjects, handles errors, and exports the first worksheet to CSV. | Write code that returns the total number of shapes removed before exporting the workbook to CSV with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, removes every drawing object (pictures, charts, text boxes) from each worksheet using `RemoveAllDrawingObjects`, then saves the first worksheet as a CSV file. Includes file‑existence validation and exception handling for reliable batch processing.
    public class RemoveShapesAndExportCsv
    {
        public static void Run(string inputPath, string outputCsvPath)
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all drawing objects from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.RemoveAllDrawingObjects();
            }

            // Save the first worksheet as CSV
            workbook.Save(outputCsvPath, SaveFormat.Csv);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Expect two arguments: input file path and output CSV path
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: RemoveShapesAndExportCsv <input.xlsx> <output.csv>");
                    return;
                }

                string inputPath = args[0];
                string outputCsvPath = args[1];

                RemoveShapesAndExportCsv.Run(inputPath, outputCsvPath);
                Console.WriteLine($"CSV file successfully saved to: {outputCsvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
