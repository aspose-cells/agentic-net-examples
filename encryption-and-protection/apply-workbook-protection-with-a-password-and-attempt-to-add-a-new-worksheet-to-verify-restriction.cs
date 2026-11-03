// Title: Protect an Excel workbook's structure with a password and verify that adding new worksheets is blocked using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect workbook structure with password and prevent new sheet creation | How to test workbook structure protection enforcement when adding worksheets in Aspose.Cells | Aspose.Cells example catching exception for adding worksheet to a password-protected workbook | Verify that workbook protection blocks worksheet addition using Aspose.Cells for .NET | Apply password protection to Excel workbook structure and check restriction with Aspose.Cells
// Tags: Aspose.Cells workbook structure protection C# | Excel password protection using Aspose.Cells | disallow new worksheet addition after protection | catch protection exception Aspose.Cells | ProtectionType.Structure usage Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, applies structure protection with a password via wb.Protect(ProtectionType.Structure, "MySecretPassword"), then attempts to add a worksheet inside a try/catch block to confirm that the protection prevents the operation. It logs the outcome and saves the protected file as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Apply workbook protection with a password (protect structure)
        wb.Protect(ProtectionType.Structure, "MySecretPassword");

        // Attempt to add a new worksheet to verify the restriction
        try
        {
            wb.Worksheets.Add("NewSheet");
            Console.WriteLine("Worksheet added successfully - protection not enforced.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to add worksheet due to protection: " + ex.Message);
        }

        // Save the workbook
        wb.Save("ProtectedWorkbook.xlsx");
    }
}
