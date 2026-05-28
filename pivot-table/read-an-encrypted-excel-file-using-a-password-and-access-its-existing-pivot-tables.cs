using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";

        // Password used to protect the workbook
        string password = "myPassword";

        try
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file \"{filePath}\" was not found.");

            // Configure load options with the password and enable pivot cache parsing
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password,
                ParsingPivotCachedRecords = true
            };

            // Load the password‑protected workbook
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Iterate through worksheets and list pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.PivotTables.Count > 0)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {sheet.PivotTables.Count} pivot table(s).");

                    for (int i = 0; i < sheet.PivotTables.Count; i++)
                    {
                        PivotTable pivot = sheet.PivotTables[i];
                        Console.WriteLine($"  PivotTable {i + 1}: Name = \"{pivot.Name}\"");
                    }
                }
            }
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (CellsException cellsEx)
        {
            // Handles errors thrown by Aspose.Cells (e.g., invalid password)
            Console.WriteLine($"Aspose.Cells error: {cellsEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}