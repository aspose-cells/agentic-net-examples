using System;
using System.IO;
using Aspose.Cells;

class RemoveWorkbookPassword
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";

        // Password used for protecting the workbook and its worksheets
        string password = "myPassword";

        // Output path for the unprotected workbook
        string outputPath = "unprotected.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook with the supplied password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove workbook protection
            workbook.Unprotect(password);

            // Remove protection from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Unprotect(password);
            }

            // Save the workbook without any password protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}