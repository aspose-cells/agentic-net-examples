using Aspose.Cells;
using System;

class RemovePasswordProtection
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "protected.xlsx";
        string outputPath = "unprotected.xlsx";

        // Password that protects the workbook (structure/window) and possibly worksheets
        string password = "myPassword";

        // Load the workbook with the password (required for encrypted files)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Remove workbook protection (structure/window) if it is set
        if (workbook.IsWorkbookProtectedWithPassword)
        {
            workbook.Unprotect(password);
        }

        // Iterate through all worksheets and remove their protection
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.IsProtected)
            {
                // Try to unprotect with the known password; if it fails, attempt parameterless unprotect
                try
                {
                    sheet.Unprotect(password);
                }
                catch
                {
                    sheet.Unprotect();
                }
            }
        }

        // Remove file encryption password (if any)
        workbook.Settings.Password = null;

        // Save the workbook without any password protection
        workbook.Save(outputPath);
    }
}