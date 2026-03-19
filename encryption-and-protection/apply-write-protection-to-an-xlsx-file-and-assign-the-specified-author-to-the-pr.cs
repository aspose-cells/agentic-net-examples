using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the write protection settings via WorkbookSettings
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Assign the author for the write protection
        writeProtection.Author = "John Doe";

        // (Optional) Set a password to activate write protection
        writeProtection.Password = "SecurePass123";

        // Save the workbook with the applied write protection
        workbook.Save("WriteProtectedWorkbook.xlsx");
    }
}