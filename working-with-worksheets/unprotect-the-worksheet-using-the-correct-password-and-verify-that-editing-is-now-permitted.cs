// Title: How to unprotect a password‑protected Excel worksheet with Aspose.Cells for .NET and confirm edit access
// Common Searches: Aspose.Cells C# unprotect worksheet with password and check IsProtected | remove worksheet protection in .NET and edit a cell after unprotecting | C# code to load protected XLSX, unprotect using password, and save unprotected file | verify cell editing is allowed after calling Worksheet.Unprotect in Aspose.Cells | example of creating a protected workbook then unprotecting it with Aspose.Cells
// Tags: worksheet unprotect Aspose.Cells C# | remove password protection from Excel sheet | verify IsProtected flag after unprotect | edit cell A1 after worksheet unprotect | save workbook as unprotected XLSX using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace UnprotectWorksheetDemo
{
    // Loads a password‑protected XLSX file, confirms the worksheet is protected, calls Worksheet.Unprotect with the correct password, rechecks the protection flag, edits cell A1 to demonstrate write access, and saves the workbook as an unprotected file.
    class Program
    {
        // Path to the input workbook (protected) and output workbook (unprotected)
        private const string InputFile = "protected.xlsx";
        private const string OutputFile = "unprotected.xlsx";
        // The password used to protect the worksheet
        private const string WorksheetPassword = "SecretPwd123";

        static void Main(string[] args)
        {
            try
            {
                // Ensure the Aspose.Cells license is set if you have one.
                // License license = new License();
                // license.SetLicense("Aspose.Cells.lic");

                // Step 1: Prepare a protected workbook if it does not already exist.
                if (!File.Exists(InputFile))
                {
                    CreateProtectedWorkbook(InputFile, WorksheetPassword);
                    Console.WriteLine($"Created protected workbook '{InputFile}'.");
                }

                // Step 2: Load the workbook.
                Workbook workbook = new Workbook(InputFile);
                Worksheet sheet = workbook.Worksheets[0];

                // Step 3: Verify that the worksheet is currently protected.
                Console.WriteLine($"Worksheet '{sheet.Name}' IsProtected before unprotect: {sheet.IsProtected}");

                // Step 4: Unprotect the worksheet using the correct password.
                sheet.Unprotect(WorksheetPassword);
                Console.WriteLine("Worksheet successfully unprotected with the correct password.");

                // Step 5: Verify that the worksheet is no longer protected.
                Console.WriteLine($"Worksheet '{sheet.Name}' IsProtected after unprotect: {sheet.IsProtected}");

                // Step 6: Attempt to edit a cell to confirm editing is permitted.
                sheet.Cells["A1"].PutValue("Edited after unprotect");
                Console.WriteLine("Cell A1 edited successfully.");

                // Step 7: Save the modified workbook.
                workbook.Save(OutputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Unprotected workbook saved as '{OutputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="path">File path to save the protected workbook.</param>
        /// <param name="password">Password to protect the worksheet.</param>
        private static void CreateProtectedWorkbook(string path, string password)
        {
            // Create a new workbook with a single worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "ProtectedSheet";

            // Write a sample value before protection.
            ws.Cells["A1"].PutValue("Original content");

            // Protect the worksheet with the provided password.
            ws.Protect(ProtectionType.All, password, null);

            // Save the workbook.
            wb.Save(path, SaveFormat.Xlsx);
        }
    }
}
