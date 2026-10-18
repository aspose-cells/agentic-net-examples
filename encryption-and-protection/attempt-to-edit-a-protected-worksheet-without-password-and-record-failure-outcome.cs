// Title: Attempt to edit a password‑protected worksheet without unprotecting it using Aspose.Cells for .NET and capture the failure
// Common Searches: Aspose.Cells C# edit cell in protected worksheet without password | How to catch CellsException when modifying a locked Excel sheet with Aspose.Cells | C# load Excel file with optional password prompt using Aspose.Cells LoadOptions | Record failure of editing a protected worksheet in Aspose.Cells example | Attempt to write to a protected sheet without unprotecting using Aspose.Cells .NET
// Tags: edit protected worksheet Aspose.Cells | handle CellsException Aspose.Cells | load workbook optional password Aspose.Cells | detect worksheet protection C# | record edit failure Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, prompting for a password if the initial load fails, accesses the first worksheet, attempts to write to cell A1 without unprotecting the sheet, catches CellsException to identify protection, and outputs whether the edit succeeded or correctly failed.
class Program
{
    static void Main()
    {
        const string filePath = "protected.xlsx";

        // Verify that the workbook file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook = null;

        // Attempt to load the workbook; if it fails, assume it may be password‑protected.
        try
        {
            workbook = new Workbook(filePath);
        }
        catch (Exception loadEx)
        {
            Console.WriteLine("Initial load failed (possible password protection): " + loadEx.Message);
            Console.Write("Enter password: ");
            string password = Console.ReadLine();

            try
            {
                LoadOptions loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception pwdEx)
            {
                Console.WriteLine("Failed to load workbook with provided password: " + pwdEx.Message);
                return;
            }
        }

        // Ensure workbook was loaded.
        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        // Access the first worksheet (index 0).
        Worksheet sheet = workbook.Worksheets[0];

        bool editSucceeded = false;

        // Try to modify a cell on the protected worksheet without unprotecting it.
        try
        {
            sheet.Cells["A1"].PutValue("Attempted Edit");
            editSucceeded = true; // Unexpected success.
        }
        catch (CellsException)
        {
            Console.WriteLine("Failed to edit protected worksheet (cell is protected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error during edit attempt: " + ex.Message);
        }

        // Output the result of the edit attempt.
        if (editSucceeded)
        {
            Console.WriteLine("Edit unexpectedly succeeded.");
        }
        else
        {
            Console.WriteLine("Edit correctly failed due to worksheet protection.");
        }
    }
}
