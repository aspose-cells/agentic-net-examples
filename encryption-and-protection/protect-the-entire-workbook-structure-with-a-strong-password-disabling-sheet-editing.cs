// Title: How to protect an Excel workbook’s structure with a strong password using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect workbook structure with password example | Set strong password for Excel workbook structure using Aspose.Cells .NET | Disable sheet adding and renaming in Excel file with Aspose.Cells C# | Workbook.Protect(ProtectionType.Structure) usage in C# Aspose.Cells | Prevent users from moving or hiding sheets in .xlsx via Aspose.Cells
// Tags: Aspose.Cells workbook structure lock C# | Excel .xlsx structure protection via Aspose | Workbook.Protect API strong password | Prevent sheet changes Aspose.Cells .NET | Structure-level protection for Excel files

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, adds worksheets, then applies workbook.Protect(ProtectionType.Structure, password) with a strong password to lock the workbook’s structure—disallowing adding, removing, renaming, moving, or hiding sheets—and saves the result as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Optional: add or rename worksheets
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets.Add("Sheet2");

            // Protect the entire workbook structure with a strong password
            // This disables adding, removing, renaming, moving, or hiding sheets
            string password = "Str0ngP@ssw0rd!2026";
            workbook.Protect(ProtectionType.Structure, password);

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Save the protected workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
