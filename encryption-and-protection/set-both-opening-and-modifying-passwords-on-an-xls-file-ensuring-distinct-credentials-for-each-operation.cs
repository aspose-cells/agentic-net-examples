// Title: How to apply separate opening and modifying passwords to an Excel 97‑2003 (.xls) file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set a read password and a different write password for an .xls workbook | protect Excel 97-2003 file with separate open and edit passwords using Aspose.Cells .NET | Aspose.Cells workbook.Settings.Password vs workbook.Protect for distinct passwords in C# | save an .xls file with both opening and modify passwords in Aspose.Cells | example code for applying read‑only and write protection passwords to an XLS file in .NET
// Tags: Aspose.Cells set opening password | Aspose.Cells write protection modify password | Workbook.Settings.Password read protection | Workbook.Protect modify password | SaveFormat.Excel97To2003 password protection

using System;
using Aspose.Cells;

// The example creates or loads a workbook, writes sample data, assigns an opening (read) password via Workbook.Settings.Password, applies write protection with a distinct modifying password using Workbook.Protect, and saves the file as an Excel 97‑2003 (.xls) workbook with both passwords applied.
class SetPasswordsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Optionally add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Set the opening (read) password
            workbook.Settings.Password = "OpenPassword123";

            // Enable write protection and set the modifying (write) password
            workbook.Protect(ProtectionType.All, "ModifyPassword456");

            // Save the workbook as an XLS file (Excel 97-2003 format)
            string outputPath = "ProtectedWorkbook.xls";
            workbook.Save(outputPath, SaveFormat.Excel97To2003);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
