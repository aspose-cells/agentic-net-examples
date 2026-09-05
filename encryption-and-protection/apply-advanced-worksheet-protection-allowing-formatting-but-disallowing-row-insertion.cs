// Title: How to protect an Aspose.Cells worksheet in .NET to allow formatting but prevent row insertion
// AI Prompts: Use Aspose.Cells in C# to protect a worksheet with a password, enable cell/column/row formatting, and block row insertion. | Configure worksheet protection options in a .NET workbook so users can format cells but cannot add new rows. | Modify an existing Excel file with Aspose.Cells to set AllowFormatting flags true and AllowInsertingRow false.
// Common Searches: Aspose.Cells C# protect worksheet allow formatting disable row insertion | How to enable formatting permissions while restricting row addition in Aspose.Cells .NET | Set worksheet protection options for formatting only using Aspose.Cells API | C# example of custom worksheet protection that blocks row insertion | Password protect Excel sheet with formatting rights but no row insert in Aspose.Cells
// Tags: worksheet protection allow formatting Aspose.Cells | disable row insertion Aspose.Cells | custom protection settings C# | password protected workbook formatting only | Aspose.Cells protection API usage

using System;
using Aspose.Cells;

// Demonstrates creating or loading a workbook, applying password protection to a worksheet, enabling cell/column/row formatting, disabling row insertion, and saving the protected Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password.
            // The third parameter is the old password; pass null when there is none.
            sheet.Protect(ProtectionType.All, "StrongPassword123", null);

            // Get the protection settings object
            Protection protection = sheet.Protection;

            // Allow formatting of cells, columns, and rows
            protection.AllowFormattingCell = true;
            protection.AllowFormattingColumn = true;
            protection.AllowFormattingRow = true;

            // Disallow insertion of rows
            protection.AllowInsertingRow = false;

            // Save the workbook with the applied protection
            workbook.Save("ProtectedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
