// Title: How to programmatically unprotect an Excel worksheet by testing common passwords using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells load password‑protected workbook with a list of possible passwords | aspnet unprotect Excel worksheet by trying multiple passwords using Aspose.Cells | iterate through common passwords to remove worksheet protection in C# | Aspose.Cells open encrypted .xlsx file with fallback passwords | C# code to unprotect worksheet when password is unknown using Aspose.Cells
// Tags: Aspose.Cells load workbook with password enumeration | Aspose.Cells worksheet unprotect multiple passwords | C# Excel worksheet password brute force | Excel .xlsx protection removal Aspose.Cells | LoadOptions password list Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example checks for a protected.xlsx file, defines a list of common passwords, attempts to open the workbook using each password via LoadOptions, then iterates through the same list to call Worksheet.Unprotect on the first sheet. If unprotected, the workbook is saved as unprotected.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "unprotected.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        // Common passwords to try for both workbook and worksheet protection
        List<string> commonPasswords = new List<string>
        {
            "", "password", "123456", "12345", "1234", "1111", "0000", "admin",
            "letmein", "welcome", "qwerty", "abc123", "iloveyou", "123123",
            "master", "sunshine", "princess", "football", "monkey", "shadow"
        };

        // Attempt to load the workbook using the password list
        Workbook workbook = LoadWorkbookWithPasswords(inputPath, commonPasswords);
        if (workbook == null)
        {
            Console.WriteLine("Failed to open the workbook. No valid password found.");
            return;
        }

        // Assume we want to unprotect the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Attempt to unprotect the worksheet
        bool success = TryUnprotectWorksheet(sheet, commonPasswords);
        Console.WriteLine(success ? "Worksheet successfully unprotected." : "Failed to unprotect worksheet.");

        // Save the workbook only if it was unprotected
        if (success)
        {
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Tries to load a workbook using a list of possible passwords
    static Workbook LoadWorkbookWithPasswords(string path, List<string> passwords)
    {
        foreach (string pwd in passwords)
        {
            try
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = pwd };
                return new Workbook(path, loadOptions);
            }
            catch (CellsException)
            {
                // Continue trying next password
            }
            catch (Exception)
            {
                // Non‑password related error; stop trying
                break;
            }
        }
        return null;
    }

    // Attempts to unprotect a worksheet using a list of possible passwords
    static bool TryUnprotectWorksheet(Worksheet sheet, List<string> passwords)
    {
        foreach (string pwd in passwords)
        {
            try
            {
                sheet.Unprotect(pwd);
                if (!sheet.IsProtected)
                {
                    return true;
                }
            }
            catch
            {
                // Ignored – continue trying next password
            }
        }
        return false;
    }
}
