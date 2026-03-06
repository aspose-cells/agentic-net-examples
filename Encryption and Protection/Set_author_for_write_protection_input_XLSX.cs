using System;
using Aspose.Cells;

class SetWriteProtectionAuthor
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the workbook's write protection settings
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Set the author name for the write protection
        writeProtection.Author = "John Doe";

        // (Optional) Set a password to activate write protection
        writeProtection.Password = "password123";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}