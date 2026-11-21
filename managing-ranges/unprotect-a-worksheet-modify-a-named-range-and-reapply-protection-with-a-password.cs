// Title: Unprotect an Excel worksheet, edit a named range, and re‑apply password protection using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unprotect worksheet modify named range and protect again | example code to change values in a defined name on a protected Excel sheet using Aspose.Cells | how to retrieve and edit a named range after unprotecting a sheet with Aspose.Cells .NET | protect Excel worksheet with password after updating range using Aspose.Cells C# | Aspose.Cells unprotect sheet without password then protect with password
// Tags: worksheet unprotect Aspose.Cells | named range edit Aspose.Cells | sheet protection password Aspose.Cells | range retrieval by name Aspose.Cells | workbook save Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing Excel file, unprotects the first worksheet, retrieves a named range, updates each cell in that range, re‑applies sheet protection with a password, and saves the modified workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";
            const string namedRangeName = "MyRange";
            const string password = "myPassword";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet (no password needed if unknown)
            sheet.Unprotect();

            // Retrieve the range by its defined name
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(namedRangeName);
            if (range != null)
            {
                // Example modification: set each cell in the range to a new value
                foreach (Cell cell in range)
                {
                    cell.PutValue("Modified");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Named range \"{namedRangeName}\" not found.");
            }

            // Re‑apply protection with a password (oldPassword is empty because the sheet is currently unprotected)
            sheet.Protect(ProtectionType.All, password, string.Empty);

            // Save the modified workbook to a new file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
