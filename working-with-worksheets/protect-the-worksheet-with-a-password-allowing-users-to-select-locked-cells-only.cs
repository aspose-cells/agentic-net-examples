// Title: Password‑protect an Excel worksheet while allowing selection of locked cells using Aspose.Cells for .NET (C#)
// Common Searches: C# code to protect an Excel worksheet with a password using Aspose.Cells and keep locked cells selectable | Aspose.Cells protect worksheet allow selection of locked cells example | How to use sheet.Protect with ProtectionType.All in Aspose.Cells for .NET | Saving a password‑protected workbook with Aspose.Cells C# | Set worksheet protection options programmatically with Aspose.Cells .NET
// Tags: Aspose.Cells worksheet password protection C# | ProtectionType.All usage Aspose.Cells | allow selection of locked cells Excel Aspose.Cells | programmatic Excel sheet protection .NET | save protected workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, applies password protection with ProtectionType.All via sheet.Protect, and saves the workbook as a password‑protected Excel file.
class ProtectWorksheetExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password (protect all options)
            // The Protect method requires the old password parameter; pass null or empty string if not needed
            string password = "MySecurePassword";
            sheet.Protect(ProtectionType.All, password, null);

            // Save the protected workbook
            string outputPath = "ProtectedWorksheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
