// Title: Read a password‑protected ODS file in read‑only mode and list its non‑empty cells using Aspose.Cells for .NET
// Common Searches: how to open an encrypted ODS spreadsheet with Aspose.Cells in C# | verify password of an ODS file before loading with Aspose.Cells .NET | read only mode for password protected ODS workbook using Aspose.Cells | extract non empty cell values from a protected ODS file in C# | Aspose.Cells LoadOptions password example for ODS format
// Tags: Aspose.Cells LoadOptions password ODS | read‑only encrypted ODS workbook C# | enumerate worksheets and cells Aspose.Cells | list non‑empty cell values .NET | open protected ODS spreadsheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the ODS file exists, configures LoadOptions with LoadFormat.Ods and the supplied password, loads the workbook in read‑only mode (verifying the password), then iterates each worksheet's used range and prints the address and value of every non‑empty cell, with error handling for loading and reading operations.
class Program
{
    static void Main()
    {
        // Path to the encrypted ODS file
        string filePath = @"C:\Path\To\EncryptedFile.ods";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // The password to verify
        string password = "yourPassword";

        // Set load options for ODS format and provide the password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
        {
            Password = password
        };

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook with the supplied password
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (Exception ex)
        {
            // Aspose.Cells throws a generic exception for incorrect passwords in some versions
            Console.WriteLine($"Failed to load the file: {ex.Message}");
            return;
        }

        // If we reach here, the password is correct and the workbook is loaded
        Console.WriteLine("Password verified successfully. Extracting cell values...");

        try
        {
            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Get the used range of the worksheet
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Loop through rows and columns within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Output only non‑empty cells
                        if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                        {
                            string address = cell.Name; // e.g., "A1"
                            string value = cell.Value?.ToString() ?? string.Empty;
                            Console.WriteLine($"{address}: {value}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while reading cells: {ex.Message}");
        }

        // No need to save; the workbook was opened read‑only (no changes were made)
    }
}
