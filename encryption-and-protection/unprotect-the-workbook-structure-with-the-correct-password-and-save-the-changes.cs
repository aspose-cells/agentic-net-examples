// Title: Remove workbook structure protection from an Excel file using a password with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens a password‑protected .xlsx, calls the Aspose.Cells Workbook.Unprotect method with the supplied password to clear structure protection, and saves the result to a new file. | Demonstrate how to confirm that the workbook is no longer protected after invoking Workbook.Unprotect and then persist the workbook using Aspose.Cells.
// Common Searches: aspnet c# how to unprotect workbook structure with password using Aspose.Cells | remove Excel workbook structure protection programmatically Aspose.Cells .NET | unprotect protected .xlsx file and save new copy C# Aspose | Workbook.Unprotect method example with password Aspose.Cells | save unprotected workbook after calling Unprotect Aspose.Cells C#
// Tags: Aspose.Cells unprotect workbook API | Excel workbook structure lock removal Aspose.Cells | save unprotected workbook C# | load password protected .xlsx Aspose.Cells | Workbook protection handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads a password‑protected Excel workbook, removes its structure protection via Workbook.Unprotect, and saves the unprotected workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string password = "YourPasswordHere";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the password
            workbook.Unprotect(password);

            // Save the workbook with the changes applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
