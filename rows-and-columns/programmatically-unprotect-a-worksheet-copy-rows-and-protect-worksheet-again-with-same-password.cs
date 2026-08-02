using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data to demonstrate row copying
        worksheet.Cells["A1"].PutValue("Row 1");
        worksheet.Cells["A2"].PutValue("Row 2");
        worksheet.Cells["A3"].PutValue("Row 3");

        // Define the protection password
        string password = "myPassword";

        // Protect the worksheet with the password
        worksheet.Protect(ProtectionType.All, password, null);

        // Unprotect the worksheet using the same password
        worksheet.Unprotect(password);

        // Copy rows 0 and 1 (first two rows) to start at row index 3 (fourth row)
        // Parameters: source cells, source start row, destination start row, number of rows to copy
        worksheet.Cells.CopyRows(worksheet.Cells, 0, 3, 2);

        // Re‑protect the worksheet with the original password
        worksheet.Protect(ProtectionType.All, password, null);

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}