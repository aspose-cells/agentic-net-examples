using System;
using Aspose.Cells;

class RemoveWorkbookProtection
{
    static void Main()
    {
        // Path to the protected workbook
        string inputPath = "protected_workbook.xlsx";

        // Path where the unprotected workbook will be saved
        string outputPath = "unprotected_workbook.xlsx";

        // Password used to protect the workbook structure (replace with actual password)
        string password = "password123";

        // Load the protected workbook
        Workbook workbook = new Workbook(inputPath);

        // Remove workbook structure protection
        workbook.Unprotect(password);

        // Save the unprotected workbook
        workbook.Save(outputPath);

        // Optional: confirm removal
        Console.WriteLine("Workbook protection removed and saved to: " + outputPath);
    }
}