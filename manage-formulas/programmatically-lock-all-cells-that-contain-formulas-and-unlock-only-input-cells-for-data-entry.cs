// Title: Lock formula cells and unlock input cells in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock cells that contain formulas and keep other cells editable | How to protect an Excel worksheet while allowing data entry in non‑formula cells using Aspose.Cells | C# code to lock all formula cells and unlock user input cells in a workbook with Aspose.Cells | Apply worksheet protection in Aspose.Cells and unlock cells without formulas | Programmatically set IsLocked property for formula cells only with Aspose.Cells .NET
// Tags: lock formula cells Aspose.Cells | unlock non‑formula cells C# | worksheet protection Aspose.Cells | apply locked style to entire sheet | styleflag locked property Aspose.Cells | protect worksheet after locking cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, applies a locked style to every cell, then iterates through each cell to unlock those that are not formulas, protects each worksheet with full protection, and saves the result to a new workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a style that locks all cells
                Style lockedStyle = workbook.CreateStyle();
                lockedStyle.IsLocked = true;

                // Apply the locked style to the entire sheet
                StyleFlag lockFlag = new StyleFlag { Locked = true };
                sheet.Cells.ApplyStyle(lockedStyle, lockFlag);

                // Unlock cells that do NOT contain formulas (input cells)
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula)
                    {
                        Style cellStyle = cell.GetStyle();
                        cellStyle.IsLocked = false;
                        cell.SetStyle(cellStyle);
                    }
                }

                // Protect the worksheet to enforce the lock settings
                sheet.Protect(ProtectionType.All);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
