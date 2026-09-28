// Title: Apply separate opening and modify passwords to an XLS workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that sets a read‑only password and a distinct edit password for an Excel 97‑2003 workbook, then saves it as .xls. | Show how to use Workbook.Settings.Password for opening protection and Workbook.Protect with ProtectionType.All for write protection in a .xls file via Aspose.Cells. | Create a sample XLS workbook, apply separate open and modify passwords, and confirm both protections are applied when saving.
// Common Searches: how to assign a read‑only password and a different edit password to an .xls workbook using Aspose.Cells in C# | Aspose.Cells example for protecting an Excel97To2003 file with separate open and write passwords | C# code to use Workbook.Settings.Password and Workbook.Protect for distinct passwords in XLS | save Excel 97‑2003 workbook with both opening and modifying passwords via Aspose.Cells .NET | apply separate open and modify passwords to an XLS file with Aspose.Cells API
// Tags: Aspose.Cells set opening password XLS | Aspose.Cells apply modify password Excel97To2003 | Workbook.Settings.Password C# example | Workbook.Protect ProtectionType.All usage | protect XLS workbook with distinct passwords

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, adds optional data, assigns a read‑only opening password via Workbook.Settings.Password, applies a separate edit password using Workbook.Protect with ProtectionType.All, and saves the file as an Excel 97‑2003 (.xls) workbook.
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

            // Set the opening password
            workbook.Settings.Password = "Open123!";

            // Apply write protection with a modifying password
            workbook.Protect(ProtectionType.All, "Modify456!");

            // Define output file path
            string outputPath = "ProtectedWorkbook.xls";

            // Save the workbook as an XLS file
            workbook.Save(outputPath, SaveFormat.Excel97To2003);

            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
