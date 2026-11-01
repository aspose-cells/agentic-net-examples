// Title: How to protect a specific worksheet with a password while still allowing cell, column, and row formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect worksheet password allow formatting changes | Enable worksheet protection but keep cell formatting editable in Aspose.Cells .NET | Set AllowFormattingCell, AllowFormattingColumn, and AllowFormattingRow on a protected sheet with Aspose.Cells | C# code to protect an Excel sheet with a password and permit formatting using Aspose.Cells | Worksheet protection with formatting permissions example for Aspose.Cells
// Tags: worksheet protection password Aspose.Cells | allow formatting on protected worksheet C# | Aspose.Cells Protect method with formatting permissions | set AllowFormattingCell property Aspose.Cells | save protected workbook Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, names the first worksheet, applies password‑based protection using ProtectionType.All, enables cell, column, and row formatting while keeping cells locked, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet with a password.
            // The overload requires oldPassword; pass null when there is none.
            sheet.Protect(ProtectionType.All, "SecurePassword123", null);

            // Allow formatting changes while keeping cells locked.
            sheet.Protection.AllowFormattingCell = true;      // allow cell formatting
            sheet.Protection.AllowFormattingColumn = true;    // allow column formatting
            sheet.Protection.AllowFormattingRow = true;       // allow row formatting

            // Determine output path and ensure directory exists
            string outputPath = "ProtectedWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
