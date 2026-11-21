// Title: How to protect an Excel worksheet while allowing only formula cells to be edited with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect worksheet but keep formula cells editable | unlock cells containing formulas before worksheet protection using Aspose.Cells | set IsLocked false for formula cells in Aspose.Cells .NET | apply password protection to worksheet while allowing formula editing Aspose.Cells
// Tags: selective worksheet protection based on formula detection Aspose.Cells | unlock cells with formulas using cell style IsLocked false C# | password-protect worksheet via Protect method Aspose.Cells | enumerate cells to identify formulas Aspose.Cells | export protected workbook to XLSX format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, scans every cell in the first worksheet, unlocks cells that contain formulas by setting their IsLocked style to false, then protects the worksheet with a password so only those formula cells remain editable, and finally saves the result as a protected XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Unlock cells that contain formulas
                Cells cells = worksheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Style style = cell.GetStyle();
                        style.IsLocked = false;
                        cell.SetStyle(style);
                    }
                }

                // Protect the worksheet; only unlocked (formula) cells remain editable
                // The third parameter is the old password (empty when not changing an existing password)
                worksheet.Protect(ProtectionType.All, "YourPassword", string.Empty);

                // Save the workbook
                string outputPath = "ProtectedWorksheet.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
