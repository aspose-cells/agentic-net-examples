// Title: Password‑protect a whole worksheet and workbook for read‑only access with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet read only with password | Set workbook level read‑only protection using Aspose.Cells .NET | How to apply password protection to an entire Excel file in C# with Aspose.Cells | Protect all cells in a worksheet and enforce read‑only mode Aspose.Cells example | C# code to protect Excel workbook and sheet with the same password using Aspose.Cells
// Tags: worksheet password protection Aspose.Cells .NET | workbook read‑only protection C# | ProtectionType.All usage Aspose.Cells | Excel file password protection C# Aspose | protect entire table Aspose.Cells API

using System;
using Aspose.Cells;

// The example creates a workbook, fills a small table, then uses ProtectionType.All to apply the same password to the worksheet and the workbook, forcing read‑only access for external users, and saves the file as ProtectedTable.xlsx.
class ProtectWorkbookExample
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // 2. Populate the workbook with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // 3. Protect the entire worksheet (read‑only for external users)
            // Provide an empty old password for compatibility with the API signature.
            sheet.Protect(ProtectionType.All, "MySecretPassword", string.Empty);

            // 4. Apply workbook‑level protection (forces read‑only mode on open)
            // Use the overload that accepts protection type and password.
            workbook.Protect(ProtectionType.All, "MySecretPassword");

            // 5. Save the workbook
            string outputPath = "ProtectedTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
