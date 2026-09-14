// Title: Unprotect an Excel workbook's structure and windows with a password using Aspose.Cells for .NET and verify sheet addition
// AI Prompts: Write C# code that opens a password‑protected Excel workbook using Aspose.Cells, calls Workbook.Unprotect with the supplied password, inserts a new worksheet, and saves the result as a different file. | Show how to remove both structure and windows protection from an Excel workbook in Aspose.Cells, then confirm that new sheets can be added without errors.
// Common Searches: asp.net how to remove workbook structure protection using Aspose.Cells | c# unprotect password protected excel file and add new sheet with Aspose | aspose.cells unprotect workbook windows and structure programmatically | example code to load protected.xlsx and save as unprotected.xlsx in C# | verify that workbook is unprotected by adding a worksheet in Aspose.Cells
// Tags: Aspose.Cells Workbook.Unprotect method | remove workbook structure lock .NET | add worksheet after unprotecting Excel file | C# load Excel file with Aspose.Cells | disable workbook windows lock Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example loads a password‑protected Excel file (protected.xlsx) with Aspose.Cells, calls Workbook.Unprotect using the given password to lift structure and windows protection, adds a new worksheet named "NewSheet" containing a test value, and saves the modified workbook as unprotected.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "protected.xlsx";
            const string outputPath = "unprotected.xlsx";
            const string password = "myPassword";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the protected workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the password
            workbook.Unprotect(password);

            // Add a new worksheet to verify that modifications are allowed
            Worksheet newSheet = workbook.Worksheets.Add("NewSheet");
            newSheet.Cells["A1"].PutValue("Sheet added after unprotecting.");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
