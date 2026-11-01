// Title: C# – Unprotect an Excel workbook’s structure and windows with a password using Aspose.Cells and save the result
// Common Searches: C# Aspose.Cells example to remove workbook structure protection with password | how to programmatically unprotect password‑protected Excel workbook using Aspose.Cells | unprotect Excel workbook windows and structure in .NET and save file | Aspose.Cells workbook.Unprotect method usage in C# | remove Excel workbook protection and save changes with Aspose.Cells
// Tags: aspose.cells workbook.unprotect method c# | remove workbook structure protection asp.net | c# unprotect password protected excel workbook | save unprotected workbook aspose.cells | excel workbook windows protection removal c#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, calls Workbook.Unprotect with the correct password to clear structure and window protection, and then saves the unprotected workbook to a new file, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string password = "YourPasswordHere";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the password
            workbook.Unprotect(password);

            // Save the workbook with the changes applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
