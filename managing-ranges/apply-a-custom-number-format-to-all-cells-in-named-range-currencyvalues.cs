// Title: Apply a custom currency number format to every cell in the named range "CurrencyValues" using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, retrieves the named range "CurrencyValues", and sets the custom number format "$#,##0.00" on each cell before saving the file. | Show how to iterate over all cells in a named range with Aspose.Cells and assign a currency style programmatically in a .NET application.
// Common Searches: aspnet aspose.cells apply custom currency format to named range | c# set number format for cells in a named range using Aspose.Cells | how to format all cells in a named range as currency with Aspose.Cells .NET | apply custom number format to range "CurrencyValues" in Excel using Aspose.Cells
// Tags: apply custom number format Aspose.Cells | named range formatting C# | currency style cells Aspose.Cells | loop through range cells .NET | save workbook after formatting Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example loads "input.xlsx", obtains the named range "CurrencyValues", loops through each cell in that range, applies the custom currency format "$#,##0.00" via cell styles, and saves the updated workbook as "output.xlsx" with appropriate error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "CurrencyValues"
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("CurrencyValues");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'CurrencyValues' not found.");
                return;
            }

            // Define the custom number format (currency with two decimals)
            const string customFormat = "$#,##0.00";

            // Apply the custom format to every cell in the named range
            Worksheet ws = namedRange.Worksheet;
            int firstRow = namedRange.FirstRow;
            int firstColumn = namedRange.FirstColumn;
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    try
                    {
                        Cell cell = ws.Cells[firstRow + i, firstColumn + j];
                        Style style = cell.GetStyle();
                        style.Custom = customFormat;
                        cell.SetStyle(style);
                    }
                    catch (Exception cellEx)
                    {
                        Console.WriteLine($"Failed to format cell at ({firstRow + i}, {firstColumn + j}): {cellEx.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
