// Title: How to protect an Excel worksheet with Aspose.Cells for .NET while allowing only formula cells to be edited
// AI Prompts: Write C# code using Aspose.Cells to lock all cells in a worksheet, unlock only those that contain formulas, and apply password protection. | Show how to mark formula cells as unlocked by modifying their style, then protect the worksheet with a password using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# protect worksheet but keep formula cells editable | unlock only formula cells before worksheet protection Aspose.Cells | set cell IsLocked false for formulas in Aspose.Cells .NET | apply password to Excel sheet while allowing formula editing using Aspose.Cells | how to iterate cells and unlock formulas in Aspose.Cells workbook
// Tags: worksheet password protection Aspose.Cells | unlock cells containing formulas Aspose.Cells C# | modify cell style to allow editing formulas | password protect sheet with formula edit access | iterate worksheet cells to set lock status Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionExample
{
    // The example creates a workbook, iterates through every cell in the first worksheet, unlocks cells that contain formulas by clearing their locked flag, then protects the worksheet with a password so only those formula cells remain editable, and finally saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Unlock cells that contain formulas so they can be edited by users
                foreach (Cell cell in worksheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        Style style = cell.GetStyle();
                        style.IsLocked = false; // Allow editing
                        cell.SetStyle(style);
                    }
                }

                // Protect the worksheet. Only unlocked cells (the formula cells) can be edited.
                // Provide the new password and an empty old password as required by the API.
                worksheet.Protect(ProtectionType.All, "myPassword", string.Empty);

                // Define output path
                string outputPath = "WorksheetProtectionExample.xlsx";

                // Ensure the directory exists (handle case where Path.GetDirectoryName returns null)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
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
}
