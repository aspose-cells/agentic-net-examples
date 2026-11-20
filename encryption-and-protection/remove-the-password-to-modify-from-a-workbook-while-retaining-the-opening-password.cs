// Title: How to remove the modify (write) password from an Excel workbook while keeping its opening password using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells clear modify password but keep file open password | remove workbook write protection without losing opening password in .NET | unprotect Excel file using Aspose.Cells while preserving the load password | how to keep opening password after calling Workbook.Unprotect in C# | Aspose.Cells example to delete modify password from protected workbook
// Tags: Aspose.Cells workbook unprotect | clear write protection .NET | retain opening password Excel | Excel password handling Aspose.Cells | clear workbook modify protection preserving open password

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook (optionally with an opening password), calls Workbook.Unprotect with an empty string to erase any modify/write password, and saves the file so the original opening password remains intact.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (add opening password to LoadOptions if required)
                var loadOptions = new LoadOptions(); // e.g., loadOptions.Password = "openPwd";
                var workbook = new Workbook(inputPath, loadOptions);

                // Remove write protection (modify password). If the workbook has a write password,
                // provide it here; an empty string clears protection when no password is set.
                workbook.Unprotect(string.Empty);

                // Save the workbook (the opening password, if any, remains unchanged)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
