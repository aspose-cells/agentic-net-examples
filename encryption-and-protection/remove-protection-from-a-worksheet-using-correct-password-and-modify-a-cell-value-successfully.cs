// Title: Remove worksheet password protection and edit a cell with Aspose.Cells for .NET (C#)
// Common Searches: C# code to unprotect an Excel worksheet using a password with Aspose.Cells | How to modify a cell after removing sheet protection in Aspose.Cells .NET | Loading a password‑protected XLSX file and saving changes using Aspose.Cells for C# | Aspose.Cells example for worksheet.Unprotect and cell update in C#
// Tags: worksheet.Unprotect method Aspose.Cells C# | load password protected XLSX Aspose.Cells | update cell value after unprotect Aspose.Cells | save modified workbook Aspose.Cells .NET | handle CellsException Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading a password‑protected XLSX workbook with Aspose.Cells, removing worksheet protection using the correct sheet password, updating cell A1, and saving the workbook as a new file while handling potential CellsException errors.
class Program
{
    static void Main()
    {
        const string inputFile = "protected.xlsx";
        const string outputFile = "unprotected_modified.xlsx";
        const string workbookPassword = "YourPasswordHere";   // password for the workbook file
        const string sheetPassword = "YourPasswordHere";      // password for the protected worksheet

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: File \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook with the required password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = workbookPassword
            };
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove protection from the worksheet using its password
            worksheet.Unprotect(sheetPassword);

            // Modify a cell value after unprotecting
            Cell targetCell = worksheet.Cells["A1"];
            targetCell.PutValue("Updated Value");

            // Save the changes to a new file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
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
