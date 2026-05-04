using System;
using Aspose.Cells;

class RemoveWorkbookProtection
{
    static void Main(string[] args)
    {
        // Path to the protected workbook
        string inputPath = "protected.xlsx";

        // Password used to protect the workbook (empty string if no password)
        string password = "yourPassword";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Unprotect the workbook structure/window with the password
        workbook.Unprotect(password);

        // Unprotect each worksheet that might be protected
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.IsProtected)
            {
                // Attempt to unprotect with the provided password; if it fails, try without a password
                try
                {
                    sheet.Unprotect(password);
                }
                catch (CellsException)
                {
                    sheet.Unprotect();
                }
            }
        }

        // Save the unprotected workbook to a new file
        string outputPath = "unprotected.xlsx";
        workbook.Save(outputPath);
    }
}