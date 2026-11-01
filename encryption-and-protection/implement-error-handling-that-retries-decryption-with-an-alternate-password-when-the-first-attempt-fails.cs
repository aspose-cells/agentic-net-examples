// Title: Implement fallback password handling to open a password‑protected Excel workbook with Aspose.Cells in C#
// Common Searches: c# Aspose.Cells load encrypted Excel file with secondary password after first password fails | how to catch wrong password exception and retry with another password using Aspose.Cells | Aspose.Cells LoadOptions password fallback example in C# | open password protected workbook using Aspose.Cells and try alternate password | C# code to attempt opening encrypted XLSX with primary then secondary password
// Tags: fallback password decryption Aspose.Cells | load encrypted workbook with alternate password C# | Aspose.Cells LoadOptions password retry | exception handling for encrypted Excel Aspose.Cells | C# try-catch secondary password Excel

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the XLSX file exists, then tries to load it with a primary password using Aspose.Cells LoadOptions. If an exception occurs (e.g., wrong password), the code catches the error and retries loading the same file with an alternate password, logging both errors when both attempts fail and confirming successful load otherwise.
class Program
{
    static void Main()
    {
        // Paths and passwords (replace with actual values or pass as arguments)
        string filePath = "input.xlsx";
        string primaryPassword = "FirstPassword";
        string alternatePassword = "SecondPassword";

        Workbook workbook = null;

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // First attempt: load workbook using the primary password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = primaryPassword
            };
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (Exception exPrimary)
        {
            // If loading fails (e.g., wrong password), try the alternate password
            try
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = alternatePassword
                };
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception exAlternate)
            {
                // Log the failure of both attempts
                Console.WriteLine("Failed to open the workbook with both passwords.");
                Console.WriteLine($"Primary attempt error: {exPrimary.Message}");
                Console.WriteLine($"Alternate attempt error: {exAlternate.Message}");
            }
        }

        // At this point 'workbook' is loaded (or null if both attempts failed)
        if (workbook != null)
        {
            // Continue processing the workbook as needed
            // Example: save a copy (using the provided save rule)
            // workbook.Save("output.xlsx");
            Console.WriteLine("Workbook loaded successfully.");
        }
    }
}
