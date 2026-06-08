using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the author for write protection
        workbook.Settings.WriteProtection.Author = "John Doe";

        // Set the password for write protection
        workbook.Settings.WriteProtection.Password = "password123";

        // (Optional) Recommend read‑only when opening
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook to a file
        workbook.Save("WriteProtectedWorkbook.xlsx");
    }
}