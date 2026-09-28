// Title: How to remove worksheet protection with a known password and update a cell using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a password‑protected XLSX file with Aspose.Cells LoadOptions, calls Worksheet.Unprotect with the correct password, changes the value of cell A1, and saves the result to a new file. | Demonstrate the sequence of loading a protected workbook, unprotecting the first worksheet, editing a cell, and persisting the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# unprotect worksheet with password and edit cell | load protected Excel file using LoadOptions.Password and save unprotected copy | remove sheet protection programmatically in .NET before updating cell values | example code to call Worksheet.Unprotect and then PutValue in Aspose.Cells
// Tags: Worksheet.Unprotect method Aspose.Cells | LoadOptions.Password protected workbook | edit cell after worksheet unprotect | save unprotected workbook Aspose.Cells | C# Aspose.Cells worksheet protection removal

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // // Loads a password‑protected Excel workbook, removes protection from the first worksheet using the supplied password, updates cell A1, and saves the modified workbook as a new unprotected file.
    class Program
    {
        static void Main()
        {
            const string inputFile = "protected.xlsx";
            const string outputFile = "unprotected_modified.xlsx";
            const string password = "myPassword";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: File \"{inputFile}\" not found.");
                    return;
                }

                // Load the protected workbook using the password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Remove worksheet protection (if any) using the same password
                worksheet.Unprotect(password);

                // Modify a cell value after unprotecting the sheet
                Cell targetCell = worksheet.Cells["A1"];
                targetCell.PutValue("Updated Value");

                // Save the workbook with the changes
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
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
}
