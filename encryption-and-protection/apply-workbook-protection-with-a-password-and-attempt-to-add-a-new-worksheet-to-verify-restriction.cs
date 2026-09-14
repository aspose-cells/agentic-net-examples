// Title: Use Aspose.Cells for .NET to password‑protect an Excel workbook and verify that adding a new worksheet is blocked
// AI Prompts: Apply full workbook protection with a password using workbook.Protect(ProtectionType.All, ...) and then try to add a worksheet, capturing any exception that occurs. | Demonstrate C# exception handling when a password‑protected workbook prevents the insertion of a new sheet with Aspose.Cells.
// Common Searches: C# Aspose.Cells protect workbook structure and windows with password and prevent new sheets | how to catch exception when adding worksheet to password protected Excel file using Aspose.Cells | Aspose.Cells workbook.Protect all protection types example C# | verify that workbook protection blocks worksheet addition in .NET
// Tags: Aspose.Cells workbook.Protect with password | prevent worksheet addition after workbook protection | C# exception handling for protected workbook sheet insertion | save password‑protected Excel file Aspose.Cells | protect workbook structure and windows Aspose.Cells

using Aspose.Cells;
using System;

// The code creates a new Workbook, applies full protection (structure and windows) with a password via workbook.Protect(ProtectionType.All, "MySecretPassword"), attempts to add a new worksheet inside a try/catch block, reports success or the protection error, and saves the file as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Apply workbook protection with a password (protect structure and windows)
        workbook.Protect(ProtectionType.All, "MySecretPassword");

        // Attempt to add a new worksheet to verify the protection restriction
        try
        {
            workbook.Worksheets.Add("NewSheet");
            Console.WriteLine("Worksheet added successfully - protection may not be enforced for adding sheets.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to add worksheet due to protection: " + ex.Message);
        }

        // Save the workbook
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}
