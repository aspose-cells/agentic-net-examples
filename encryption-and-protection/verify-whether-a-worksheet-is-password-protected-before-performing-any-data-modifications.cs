// Title: Check worksheet protection status and unprotect an Excel sheet before editing with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to determine if a worksheet is password protected | Aspose.Cells unprotect Excel worksheet programmatically before modifying cells | Detect worksheet protection status using Aspose.Cells .NET API | Remove password protection from a sheet with Aspose.Cells C# example | Verify and unprotect protected worksheet in Aspose.Cells before saving
// Tags: worksheet protection detection Aspose.Cells | worksheet unprotect method Aspose.Cells | edit protected Excel sheet C# Aspose.Cells | worksheet password status query .NET | Aspose.Cells unprotect operation example

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, checks if the first worksheet is password‑protected, attempts to unprotect it with a supplied password (ignoring errors if it is already unprotected or the password is incorrect), updates cell A1, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string password = "yourPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Attempt to unprotect the worksheet (if it is protected)
            try
            {
                // Unprotect using the password (if required). If the sheet is not protected, this call has no effect.
                sheet.Unprotect(password);
                Console.WriteLine("Worksheet unprotected (if it was protected).");
            }
            catch (Exception ex)
            {
                // Log any issues (e.g., wrong password) and continue
                Console.WriteLine($"Unprotect operation skipped or failed: {ex.Message}");
            }

            // Perform data modifications
            sheet.Cells["A1"].PutValue("Modified data");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
