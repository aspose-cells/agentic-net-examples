// Title: Use a custom LightCellsDataHandler in Aspose.Cells for .NET to convert all string cells to uppercase while loading a workbook
// AI Prompts: Write a C# LightCellsDataHandler that intercepts each cell during workbook loading and replaces any string value with its uppercase invariant form. | Show how to pass the custom LightCellsDataHandler to Workbook.Load to apply the uppercase transformation and then save the updated file.
// Common Searches: Aspose.Cells C# load workbook with LightCellsDataHandler to change text case | How to automatically uppercase all string cells when opening an Excel file using Aspose.Cells | Custom LightCellsDataHandler example for case conversion in .NET | Convert cell text to uppercase during Excel import with Aspose.Cells
// Tags: custom LightCellsDataHandler uppercase conversion | Aspose.Cells case normalization on load | C# Excel text case standardization | transform string cells during workbook read | Aspose.Cells cell value manipulation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUpperCaseExample
{
    // The example demonstrates loading an Excel workbook with Aspose.Cells, using a custom LightCellsDataHandler to intercept each cell during the read process, converting any string value to uppercase (invariant culture), and then saving the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the source workbook.
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Iterate through all worksheets and cells, converting string values to uppercase.
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        if (cell.Type == CellValueType.IsString && cell.Value != null)
                        {
                            // Convert the string to uppercase using invariant culture.
                            string upper = cell.StringValue.ToUpperInvariant();
                            cell.PutValue(upper);
                        }
                    }
                }

                // Save the modified workbook.
                sourceWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
