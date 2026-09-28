// Title: Unprotect an Excel worksheet, modify a named range, and re‑apply password protection with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to load an .xlsx file, call Unprotect on a worksheet, change the first cell of a named range, then protect the sheet with a password. | Show how to retrieve a named range by name, update its cell values, and re‑enable full worksheet protection with a custom password in Aspose.Cells. | Provide a step‑by‑step example that unprotects a protected sheet, edits a defined name, and saves the workbook while preserving protection settings.
// Common Searches: Aspose.Cells C# unprotect worksheet, edit named range, and protect again | how to change a named range value in a protected Excel file using Aspose.Cells | C# code to remove worksheet protection, update MyRange, and set password with Aspose.Cells | example of worksheet.Protect with password after modifying named range in Aspose.Cells | unprotect and protect Excel sheet programmatically with Aspose.Cells .NET
// Tags: worksheet unprotect Aspose.Cells C# | named range update Aspose.Cells | worksheet protect password Aspose.Cells | Aspose.Cells modify protected sheet | C# Aspose.Cells named range edit

using System;
using System.IO;
using Aspose.Cells;

// // Loads 'input.xlsx', removes protection from the first worksheet, updates the first cell of the named range 'MyRange' to "Updated Value", reapplies full protection with password "MySecurePassword", and saves the result as 'output.xlsx' using Aspose.Cells for .NET.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Unprotect the worksheet (if it is protected)
            // -------------------------------------------------
            worksheet.Unprotect();

            // -------------------------------------------------
            // 2. Modify a named range
            // -------------------------------------------------
            // Retrieve the named range by its name (e.g., "MyRange")
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange != null)
            {
                // Create a Range object that points to the cells referenced by the named range
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(namedRange.RefersTo);

                // Example modification: set a new value in the first cell of the range
                range[0, 0].PutValue("Updated Value");
            }
            else
            {
                Console.WriteLine("Named range 'MyRange' not found.");
            }

            // -------------------------------------------------
            // 3. Re‑apply protection with a password
            // -------------------------------------------------
            // Protect the worksheet with all protection types and a password
            worksheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

            // -------------------------------------------------
            // Save the modified workbook
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
