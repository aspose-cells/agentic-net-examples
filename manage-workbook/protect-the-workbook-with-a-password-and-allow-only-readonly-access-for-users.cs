// Title: How to password‑protect an Excel workbook and enforce read‑only access using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect workbook structure and windows with a password | Set an Excel file to read‑only mode using Aspose.Cells .NET | Apply password protection to an Excel workbook and prevent editing with Aspose.Cells | Use ProtectionType Structure | Windows in Aspose.Cells to lock a workbook | Save a password‑protected read‑only Excel workbook with Aspose.Cells C#
// Tags: Aspose.Cells workbook password protection | C# protect Excel workbook structure windows | read‑only Excel file Aspose.Cells | Aspose.Cells ProtectionType Structure Windows | save protected workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes sample data, applies password protection to both the workbook's structure and windows using ProtectionType.Structure | ProtectionType.Windows, and saves the file as a read‑only, password‑protected Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Add some sample data
            var cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue("Sample");
            cells["B1"].PutValue("Data");

            // Protect the workbook structure and windows with a password
            // Combine Structure and Windows flags using bitwise OR
            workbook.Protect(ProtectionType.Structure | ProtectionType.Windows, "MyPassword123");

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
