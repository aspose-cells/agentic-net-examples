// Title: How to protect an Excel workbook’s structure and disable sheet editing with a strong password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect workbook structure with password | C# code to disable editing of all worksheets in an Excel file using Aspose.Cells | set a strong password for workbook and worksheet protection in Aspose.Cells .NET | prevent adding, deleting, renaming sheets in Excel via Aspose.Cells protect method | apply full worksheet protection programmatically with Aspose.Cells in C#
// Tags: Aspose.Cells workbook structure protection | worksheet full protection Aspose.Cells | strong password Excel protection .NET | prevent sheet modifications Aspose.Cells | C# protect Excel workbook with password

using Aspose.Cells;
using System;

// The example creates a new Workbook, adds optional data, defines a strong password, protects the workbook's structure to block sheet addition, deletion, and renaming, applies full protection to each worksheet to stop cell editing, and saves the file as ProtectedWorkbook.xlsx, with error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Strong password for protection
            string password = "Str0ngP@ssw0rd!2026";

            // Protect the workbook structure (prevent adding, deleting, renaming sheets)
            workbook.Protect(ProtectionType.Structure, password);

            // Protect each worksheet to disable cell editing
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // The third parameter is the old password (not needed here), pass null or empty string
                ws.Protect(ProtectionType.All, password, null);
            }

            // Save the protected workbook
            workbook.Save("ProtectedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
