// Title: How to unprotect an Excel worksheet, retain its formatting permissions, and modify a cell formula using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unprotect worksheet while keeping AllowFormattingCell permission | Update cell formula after re‑protecting an Excel sheet with Aspose.Cells .NET | Preserve worksheet formatting permissions when changing protection in Aspose.Cells | C# code to unprotect, edit, and protect an Excel worksheet with password using Aspose.Cells | Restore AllowFormattingColumn and AllowFormattingRow after modifying a protected sheet in Aspose.Cells
// Tags: worksheet unprotect preserve formatting Aspose.Cells | update cell formula protected worksheet C# | Aspose.Cells worksheet protection API | restore AllowFormattingRow AllowFormattingColumn Aspose.Cells | excel password protect unprotect Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, saves the current AllowFormattingCell/Column/Row settings, unprotects the first worksheet with a password, changes the formula of cell B2 to "=SUM(A1:A10)", re‑protects the sheet using the same password, restores the saved formatting permissions, and saves the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "password";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Preserve current formatting permissions before unprotecting
            bool allowFormattingCell = worksheet.Protection.AllowFormattingCell;
            bool allowFormattingColumn = worksheet.Protection.AllowFormattingColumn;
            bool allowFormattingRow = worksheet.Protection.AllowFormattingRow;

            // Unprotect the worksheet (provide the password if one was set)
            worksheet.Unprotect(password);

            // Update the formula of a specific cell (e.g., B2)
            Cell targetCell = worksheet.Cells["B2"];
            targetCell.Formula = "=SUM(A1:A10)";

            // Re‑protect the worksheet with the same password
            // The overload requires oldPassword; pass the same password (or null if not needed)
            worksheet.Protect(ProtectionType.All, password, password);

            // Restore the previously saved formatting permissions
            worksheet.Protection.AllowFormattingCell = allowFormattingCell;
            worksheet.Protection.AllowFormattingColumn = allowFormattingColumn;
            worksheet.Protection.AllowFormattingRow = allowFormattingRow;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
