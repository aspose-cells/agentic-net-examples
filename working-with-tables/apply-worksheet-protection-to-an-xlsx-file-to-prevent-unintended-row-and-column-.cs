using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the protection settings for the worksheet
        Protection protection = worksheet.Protection;

        // Disallow deletion of rows and columns
        protection.AllowDeletingRow = false;
        protection.AllowDeletingColumn = false;

        // Apply protection to the worksheet (no password, protect all aspects)
        worksheet.Protect(ProtectionType.All);

        // Save the protected workbook to a new file
        workbook.Save("protected_output.xlsx");
    }
}