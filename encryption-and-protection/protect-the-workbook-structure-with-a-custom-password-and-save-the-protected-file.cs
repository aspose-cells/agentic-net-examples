// Title: Protect Excel workbook structure with a custom password in C# using Aspose.Cells and save as XLSX
// Common Searches: C# Aspose.Cells protect workbook structure with password and save as xlsx | How to lock Excel workbook structure programmatically using Aspose.Cells .NET | Set custom password for workbook protection in Aspose.Cells for .NET | Saving a password protected Excel file with Aspose.Cells C# | Aspose.Cells workbook.Protect structure example in C#
// Tags: Aspose.Cells workbook.Protect method | Excel workbook structure lock C# | save password‑protected XLSX Aspose.Cells | custom password for workbook protection .NET | protect workbook structure Aspose.Cells API

using Aspose.Cells;
using System;

// The example creates a new Workbook, applies structure protection with a custom password via workbook.Protect(ProtectionType.Structure, password), and saves the protected file as 'ProtectedWorkbook.xlsx', handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a custom password
            string password = "MySecretPassword";
            workbook.Protect(ProtectionType.Structure, password);

            // Save the protected workbook to a file
            workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
