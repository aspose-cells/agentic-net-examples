// Title: Protect a single worksheet with a password in Aspose.Cells for .NET while allowing cell, row, and column formatting
// AI Prompts: Write C# code using Aspose.Cells to apply password protection to a specific worksheet and enable formatting of cells, rows, and columns. | Show how to configure the Protection object (AllowFormattingCell, AllowFormattingRow, AllowFormattingColumn) for a worksheet in Aspose.Cells. | Create an example that creates a workbook, renames the first sheet, protects it with a password, and saves the file as an .xlsx.
// Common Searches: Aspose.Cells C# protect worksheet with password but keep cell formatting enabled | How to allow row and column formatting on a locked Excel sheet using Aspose.Cells | Set worksheet protection options programmatically in Aspose.Cells for .NET | Example of protecting a single sheet while permitting formatting changes in Aspose.Cells
// Tags: worksheet password protection Aspose.Cells | allow formatting on protected sheet C# | Aspose.Cells Protection object configuration | protect single Excel worksheet Aspose.Cells | C# Aspose.Cells worksheet protection example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, renames the first worksheet, applies password protection to that sheet, enables cell, row, and column formatting while restricting other edits, ensures the output directory exists, and saves the workbook as ProtectedWorkbook.xlsx.
class WorksheetProtectionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet with a password (oldPassword is empty for new protection)
            sheet.Protect(ProtectionType.All, "MySecretPassword", string.Empty);

            // Configure protection options (use correct property names)
            Protection protection = sheet.Protection;
            protection.AllowFormattingCell = true;    // Allow formatting cells
            protection.AllowFormattingRow = true;     // Allow formatting rows
            protection.AllowFormattingColumn = true;  // Allow formatting columns

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
