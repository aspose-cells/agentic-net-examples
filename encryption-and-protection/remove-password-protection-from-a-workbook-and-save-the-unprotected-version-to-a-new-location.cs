// Title: Remove an opening password from an Excel workbook and save an unprotected copy using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load password‑protected .xlsx and export without password | How to programmatically delete workbook opening password with Aspose.Cells | C# example to open encrypted Excel file using LoadOptions.Password and save plain version | Remove workbook protection and write new file with Aspose.Cells .NET | Save unprotected copy of a protected Excel workbook using Aspose.Cells API
// Tags: remove opening password Aspose.Cells | load protected Excel workbook C# | clear Workbook.Settings.Password | save workbook without password Aspose.Cells | Aspose.Cells password removal workflow

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source Excel file exists, creates the output folder if needed, loads the password‑protected workbook using LoadOptions.Password, clears the opening password by setting Workbook.Settings.Password to null, and saves the workbook as a new unprotected .xlsx file, handling any errors that may occur.
class RemoveWorkbookPassword
{
    static void Main()
    {
        try
        {
            // Paths for the source (protected) and destination (unprotected) workbooks
            string sourcePath = @"C:\Input\ProtectedWorkbook.xlsx";
            string destinationPath = @"C:\Output\UnprotectedWorkbook.xlsx";

            // Password used to open the protected workbook
            string openingPassword = "YourPassword";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook with the opening password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = openingPassword
            };

            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Remove the opening password
            workbook.Settings.Password = null;

            // NOTE: Workbook-level protection (structure/windows) removal is omitted
            // because the Workbook.Protection property may not be available in all versions.
            // If needed, use workbook.Protection methods in a version that supports it.

            // Save the workbook without any password protection
            workbook.Save(destinationPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved without password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
