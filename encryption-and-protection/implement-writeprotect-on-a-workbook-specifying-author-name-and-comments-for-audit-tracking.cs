// Title: Apply password write protection, set author, and add audit comments to an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect Excel workbook with password and set author name | How to add custom document property for audit comments using Aspose.Cells .NET | Aspose.Cells example for write‑protecting a workbook and recording audit metadata | Saving a password‑protected .xlsx file with author and comments via Aspose.Cells | Set workbook author and custom property before saving protected Excel file in C#
// Tags: workbook write protection Aspose.Cells | set workbook author Aspose.Cells .NET | custom document property audit Aspose.Cells | protect Excel file with password Aspose.Cells | save protected workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program creates a new workbook, inserts sample data, applies write protection with a password, assigns an author name, adds a custom document property for audit comments, and saves the result as a password‑protected Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Apply write protection with a password
            workbook.Protect(ProtectionType.All, "StrongPassword123");

            // Set author name
            workbook.Settings.Author = "John Doe";

            // Add a custom document property for audit comments
            string auditComment = $"Workbook protected for audit purposes on {DateTime.Now:yyyy-MM-dd}";
            workbook.CustomDocumentProperties.Add("Comments", auditComment);

            // Define output path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
