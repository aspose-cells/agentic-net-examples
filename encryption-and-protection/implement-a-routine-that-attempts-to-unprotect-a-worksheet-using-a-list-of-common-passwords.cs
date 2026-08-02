// Title: C# Aspose.Cells Routine to Unprotect an Excel Worksheet Using a Common Password List
// Description: Loads a workbook with Aspose.Cells, checks if the first worksheet is protected, then iterates through a predefined list of common passwords calling Worksheet.Unprotect. On success the workbook is saved without protection; otherwise the original file is saved and the result is logged.
// Keywords: Aspose.Cells unprotect worksheet C# | Excel worksheet password brute force | C# try common passwords Aspose | remove worksheet protection programmatically | Worksheet.Unprotect method example | batch unprotect Excel sheets | Aspose.Cells security automation
// Common Searches: how to unprotect an Excel worksheet with Aspose.Cells in C# | C# code to try multiple passwords on a protected worksheet | Aspose.Cells worksheet protection removal script | automate Excel sheet password cracking using Aspose | unprotect Excel worksheet without original password C#
// Developer Intent: Programmatically test a set of common passwords to unlock a protected worksheet.
// Use Cases: Recover data from legacy workbooks that use simple worksheet passwords before performing analysis. | Batch‑process a directory of protected sheets, attempting common passwords to enable further automation. | Integrate into a migration tool that must remove worksheet protection prior to applying schema changes.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over a custom password file and unprotect a worksheet, handling errors and stopping after the first successful password. | Provide an optimized version of the unprotect routine that logs the password used for each worksheet and skips sheets that are already unprotected. | Suggest enhancements to parallelize password attempts across multiple worksheets while respecting Aspose.Cells thread‑safety guidelines.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook with Aspose.Cells, checks if the first worksheet is protected, then iterates through a predefined list of common passwords calling Worksheet.Unprotect. On success the workbook is saved without protection; otherwise the original file is saved and the result is logged.
public class WorksheetUnprotectHelper
{
    // List of common passwords to try.
    private static readonly List<string> CommonPasswords = new List<string>
    {
        "password",
        "123456",
        "admin",
        "test",
        "1234",
        "abcd",
        "secret",
        "letmein",
        "welcome",
        "qwerty"
    };

    /// <param name="inputFilePath">Path to the protected workbook.</param>
    /// <param name="outputFilePath">Path where the unprotected workbook will be saved.</param>
    public static void UnprotectWorksheetWithCommonPasswords(string inputFilePath, string outputFilePath)
    {
        // Load the workbook (no password is supplied because we are dealing with worksheet protection, not file encryption).
        Workbook workbook = new Workbook(inputFilePath);

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // If the worksheet is not protected, simply save and exit.
        if (!worksheet.IsProtected)
        {
            Console.WriteLine("Worksheet is not protected. Saving without changes.");
            workbook.Save(outputFilePath);
            return;
        }

        bool unprotected = false;

        // Try each password in the list.
        foreach (string pwd in CommonPasswords)
        {
            try
            {
                worksheet.Unprotect(pwd);
                if (!worksheet.IsProtected)
                {
                    Console.WriteLine($"Worksheet successfully unprotected with password: \"{pwd}\"");
                    unprotected = true;
                    break;
                }
            }
            catch (Exception ex)
            {
                // Unprotect throws an exception when the password is incorrect.
                Console.WriteLine($"Password \"{pwd}\" failed: {ex.Message}");
            }
        }

        if (!unprotected)
        {
            Console.WriteLine("Failed to unprotect the worksheet with the provided common passwords.");
        }

        // Save the workbook (whether unprotected or not).
        workbook.Save(outputFilePath);
    }

    // Example usage.
    public static void Main()
    {
        string inputPath = "ProtectedWorksheet.xlsx";
        string outputPath = "UnprotectedWorksheet.xlsx";

        UnprotectWorksheetWithCommonPasswords(inputPath, outputPath);
    }
}
