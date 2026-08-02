using System;
using Aspose.Cells;

class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add worksheets to the workbook (example: total 12 worksheets)
        for (int i = 0; i < 12; i++)
        {
            workbook.Worksheets.Add();
        }

        // Apply password protection only if there are more than ten worksheets
        if (workbook.Worksheets.Count > 10)
        {
            // Protect the workbook's structure with a password
            workbook.Protect(ProtectionType.Structure, "StrongPassword!123");
        }

        // Save the workbook to a file
        workbook.Save("ProtectedIfMoreThanTen.xlsx", SaveFormat.Xlsx);

        // Clean up resources
        workbook.Dispose();
    }
}