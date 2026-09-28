// Title: Copy rows from a password‑protected Excel worksheet by temporarily disabling protection using Aspose.Cells for .NET
// AI Prompts: Write C# code that checks whether a worksheet has a password, removes its protection, copies a specified block of rows to another sheet, and then restores the original protection settings with the same password via Aspose.Cells. | Create a reusable C# method that takes source and destination worksheet names, start row, and row count, automatically handles unprotecting and re‑protecting the source sheet while copying rows using Aspose.Cells.
// Common Searches: Aspose.Cells how to copy rows from a protected worksheet in C# | temporarily unprotect Excel sheet to copy rows using Aspose.Cells .NET | preserve worksheet password after copying rows with Aspose.Cells | C# copy rows between worksheets while keeping sheet protection | unprotect and protect worksheet programmatically Aspose.Cells example
// Tags: copy rows Aspose.Cells .NET | unprotect worksheet Aspose.Cells | protect worksheet after copy Aspose.Cells | password protected Excel sheet handling Aspose.Cells | copy rows between worksheets C# Aspose.Cells | temporary worksheet unprotection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, temporarily removes protection from the first worksheet (handling both password‑protected and unprotected cases), copies a defined range of rows to a newly added worksheet, then reapplies the original protection settings and saves the updated file.
    public class CopyRowsFromProtectedWorksheet
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook containing the protected worksheet
                Workbook workbook = new Workbook(inputPath);

                // Source (protected) worksheet
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Destination worksheet for copied rows
                Worksheet destSheet = workbook.Worksheets.Add("CopiedRows");

                // Preserve original protection settings
                Protection originalProtection = sourceSheet.Protection;

                // Determine if the worksheet is password‑protected
                bool isPasswordProtected = !string.IsNullOrEmpty(originalProtection.Password);
                string password = originalProtection.Password;

                // Unprotect the source worksheet
                if (isPasswordProtected)
                    sourceSheet.Unprotect(password);
                else
                    sourceSheet.Unprotect();

                // Define rows to copy (example: rows 0‑4)
                int sourceStartRow = 0;   // zero‑based index
                int rowsToCopy = 5;       // number of rows
                int destStartRow = 0;     // destination start row

                // Copy rows from source to destination
                destSheet.Cells.CopyRows(sourceSheet.Cells, sourceStartRow, destStartRow, rowsToCopy);

                // Re‑apply protection to the source worksheet
                if (isPasswordProtected)
                    sourceSheet.Protect(ProtectionType.All, password, null);
                else
                    sourceSheet.Protect(ProtectionType.All);

                // Restore any additional protection options
                sourceSheet.Protection.Copy(originalProtection);

                // Save the workbook with the copied rows
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyRowsFromProtectedWorksheet.Run();
        }
    }
}
