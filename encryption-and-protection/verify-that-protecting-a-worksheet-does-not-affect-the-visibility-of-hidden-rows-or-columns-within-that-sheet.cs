// Title: Verify that protecting a worksheet with Aspose.Cells for .NET retains hidden rows and columns
// AI Prompts: Write C# code using Aspose.Cells to hide a specific row and column, apply full worksheet protection with a password, then read the IsHidden properties to confirm they stay true. | Adapt an existing Aspose.Cells workbook to protect the first worksheet while preserving the hidden state of rows and columns, and output the verification results. | Generate a C# example that demonstrates checking hidden row and column status after calling Worksheet.Protect with ProtectionType.All in Aspose.Cells.
// Common Searches: Aspose.Cells C# protect worksheet keep hidden rows | does worksheet protection affect hidden columns in Aspose.Cells | how to verify hidden status after protecting an Excel sheet with Aspose.Cells | C# Aspose.Cells preserve hidden rows when applying worksheet protection | check IsHidden flag after Worksheet.Protect in Aspose.Cells .NET
// Tags: Aspose.Cells worksheet protection hidden rows | C# Aspose.Cells preserve hidden columns | Worksheet.Protect retains IsHidden property | Excel file protection without visibility change Aspose | Aspose.Cells verify hidden elements after protect

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, hides row 2 and column B, protects the first worksheet with all protection types using a password, then reads the IsHidden properties of the hidden row and column to confirm they remain true, prints the results, and saves the file as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide row 2 (index 1) and column B (index 1)
            sheet.Cells.Rows[1].IsHidden = true;
            sheet.Cells.Columns[1].IsHidden = true;

            // Protect the worksheet with a password (all protection types)
            sheet.Protect(ProtectionType.All, "password123", string.Empty);

            // Verify hidden status after protection
            bool isRowHidden = sheet.Cells.Rows[1].IsHidden;
            bool isColumnHidden = sheet.Cells.Columns[1].IsHidden;

            Console.WriteLine($"Row 2 hidden after protection: {isRowHidden}");
            Console.WriteLine($"Column B hidden after protection: {isColumnHidden}");

            // Save the workbook
            string outputPath = "ProtectedWorksheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
