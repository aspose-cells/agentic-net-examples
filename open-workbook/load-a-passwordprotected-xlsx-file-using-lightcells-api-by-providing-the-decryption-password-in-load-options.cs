// Title: Opening a password‑protected XLSX workbook with Aspose.Cells LightCells API in C# using LoadOptions.Password
// Common Searches: C# Aspose.Cells LightCells load password protected xlsx with LoadOptions | Load encrypted Excel workbook using Aspose.Cells LightCells API and password | Read protected XLSX file without user prompt using Aspose.Cells LoadOptions.Password | How to decrypt an XLSX file with Aspose.Cells LightCells in .NET | Open password protected Excel file and save unprotected copy using Aspose.Cells C#
// Tags: Aspose.Cells LightCells load password protected XLSX | LoadOptions.Password for encrypted Excel files | C# read protected workbook with LightCells | Decrypt XLSX using Aspose.Cells LoadOptions | Save unprotected copy after loading encrypted workbook

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to open a password‑protected XLSX file with Aspose.Cells LightCells API in C#. It creates a LoadOptions object with the decryption password, loads the workbook, iterates through used cells to display their values, and optionally saves an unprotected copy, handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the password‑protected XLSX file
            string inputPath = "protected.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Decryption password
            string password = "myPassword";

            // Load the workbook with the password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            var workbook = new Workbook(inputPath, loadOptions);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through used cells and display their values
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];
                    if (cell != null && cell.Value != null)
                    {
                        Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                    }
                }
            }

            // Save the workbook to a new file (optional)
            string outputPath = "unprotected.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
