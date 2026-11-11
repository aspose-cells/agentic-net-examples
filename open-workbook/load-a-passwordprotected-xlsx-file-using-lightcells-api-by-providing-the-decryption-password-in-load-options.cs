// Title: Open a password‑protected XLSX workbook with Aspose.Cells LightCells API in C# using LoadOptions.Password
// Common Searches: C# how to load an encrypted XLSX file with Aspose.Cells LightCells API | Aspose.Cells LoadOptions password property example for protected Excel workbook | Read data from a password‑protected Excel file using LightCells in .NET | Open and save a secured XLSX workbook with Aspose.Cells for .NET | LoadOptions.Password usage for encrypted workbook in Aspose.Cells C#
// Tags: Aspose.Cells LightCells load encrypted XLSX | LoadOptions password property C# | read cells from protected Excel workbook | open password protected workbook .NET | save copy of secured workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading a password‑protected XLSX file with Aspose.Cells LightCells API in C# by setting LoadOptions.Password, iterating through the first 100 rows and 20 columns to print cell values, and saving a copy of the workbook.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLSX file
        string inputFile = "protected.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        // Decryption password
        string password = "myPassword";

        try
        {
            // Set load options with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the workbook using the password
            Workbook workbook = new Workbook(inputFile, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range to read (first 100 rows, first 20 columns)
            int rowsToRead = Math.Min(100, cells.MaxDataRow + 1);
            int colsToRead = Math.Min(20, cells.MaxDataColumn + 1);

            // Process the retrieved cells
            for (int row = 0; row < rowsToRead; row++)
            {
                for (int col = 0; col < colsToRead; col++)
                {
                    Console.WriteLine(cells[row, col].Value);
                }
            }

            // Example: save a copy of the workbook (still password‑protected)
            string outputFile = "copy.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
