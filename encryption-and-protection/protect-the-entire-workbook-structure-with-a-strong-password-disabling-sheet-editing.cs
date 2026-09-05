// Title: Protect workbook structure and all worksheets with a strong password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing Excel file (or creates a new workbook), sets a strong password, protects the workbook structure, applies full protection to every worksheet, and saves the result as an XLSX using Aspose.Cells. | Show how to use Aspose.Cells ProtectionType.Structure and ProtectionType.All to enforce workbook‑level and sheet‑level password protection in a .NET application. | Demonstrate proper error handling while applying password protection to a workbook and saving the protected file with Aspose.Cells.
// Common Searches: Aspose.Cells C# protect entire workbook structure with password | How to apply password protection to all sheets in an Excel file using Aspose.Cells .NET | Set strong password for workbook and sheet protection in Aspose.Cells example | Prevent adding or deleting worksheets programmatically with Aspose.Cells
// Tags: Aspose.Cells workbook-level protection API | Aspose.Cells sheet-level protection API | Aspose.Cells generate encrypted XLSX | Aspose.Cells disable sheet modifications .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading or creating a workbook, defining a strong password, protecting the workbook structure, applying full protection to each worksheet, and saving the file as a password‑protected XLSX using Aspose.Cells for .NET.
class WorkbookProtectionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if the file exists)
            Workbook workbook;
            string inputPath = "input.xlsx";

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new blank workbook
            }

            // Define a strong password
            string password = "Str0ngP@ssw0rd!2026";

            // Protect the workbook structure to prevent adding, deleting, or renaming sheets
            workbook.Protect(ProtectionType.Structure, password);

            // Disable editing on all worksheets by applying full protection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // The third parameter is the old password (empty because the sheet is not yet protected)
                sheet.Protect(ProtectionType.All, password, string.Empty);
            }

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
