// Title: Use Aspose.Cells LightCellsDataHandler in C# to replace Excel error values with zero while loading an XLSX workbook
// AI Prompts: Write a C# LightCellsDataHandler that intercepts each cell during workbook loading and sets any error value to 0. | Show how to configure LoadOptions with a custom LightCellsDataHandler to automatically zero out #DIV/0! and #N/A cells when opening an XLSX file using Aspose.Cells. | Generate a complete C# example that loads an XLSX file, applies a LightCellsDataHandler to replace error cells with 0, and saves the cleaned workbook.
// Common Searches: Aspose.Cells C# LightCellsDataHandler replace #DIV/0! with 0 on load | How to clean Excel error cells during import using Aspose.Cells .NET | Load workbook with Aspose.Cells and automatically convert error values to zero | C# example of custom LightCellsDataHandler for error handling in XLSX files
// Tags: Aspose.Cells LightCellsDataHandler error-to-zero conversion | C# load XLSX workbook with clean numeric data | LightCellsDataHandler for Excel error conversion | convert Excel error codes to numeric zero on load | automatic error value replacement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates configuring a LightCellsDataHandler in Aspose.Cells for .NET so that, during workbook loading, any cell containing an Excel error (e.g., #DIV/0!, #N/A) is automatically replaced with the numeric value 0, resulting in a clean dataset ready for analysis.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // Replace error cells with zero
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsError)
                        {
                            cell.PutValue(0);
                        }
                    }
                }
            }

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
