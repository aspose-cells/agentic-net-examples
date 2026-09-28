// Title: Lock all cells in an Excel worksheet except a specific named range using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program with Aspose.Cells that locks every cell in a worksheet, then unlocks the cells belonging to a given named range and protects the sheet. | Show how to load an XLSX file, apply a locked style to the whole sheet, remove the lock from a named range, and save the workbook using Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells lock entire sheet but keep MyRange editable | c# protect Excel worksheet while allowing edits in a named range using Aspose.Cells | how to apply locked style to all cells and unlock a specific named range with Aspose.Cells | retrieve and modify named range protection Aspose.Cells C#
// Tags: lock cells worksheet Aspose.Cells | unlock named range Aspose.Cells | protect worksheet with editable named range | apply style to entire worksheet C# | retrieve named range by name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx, retrieves the named range "MyRange", applies a locked style to every cell, then applies an unlocked style to the cells inside that range, protects the worksheet with full protection, and saves the result as output.xlsx, handling missing files or missing named ranges gracefully.
class LockCellsOutsideNamedRange
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range (replace "MyRange" with your actual named range name)
            Aspose.Cells.Range namedRange = null;
            try
            {
                // Named ranges are stored in the Names collection; access by name
                Name name = workbook.Worksheets.Names["MyRange"];
                if (name != null)
                {
                    namedRange = name.GetRange();
                }
                else
                {
                    Console.WriteLine("Warning: Named range 'MyRange' does not exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to retrieve named range 'MyRange'. {ex.Message}");
            }

            // --------------------------------------------------------------------
            // 1. Lock all cells in the worksheet
            // --------------------------------------------------------------------
            Style lockStyle = workbook.CreateStyle();
            lockStyle.IsLocked = true;

            // Apply the lock style to all cells
            StyleFlag lockFlag = new StyleFlag { All = true };
            sheet.Cells.ApplyStyle(lockStyle, lockFlag);

            // --------------------------------------------------------------------
            // 2. Unlock cells inside the named range (if it exists)
            // --------------------------------------------------------------------
            if (namedRange != null)
            {
                Style unlockStyle = workbook.CreateStyle();
                unlockStyle.IsLocked = false;

                // Apply the unlocked style to the named range
                namedRange.ApplyStyle(unlockStyle, lockFlag);
            }

            // --------------------------------------------------------------------
            // 3. Protect the worksheet so that locked cells cannot be edited
            // --------------------------------------------------------------------
            sheet.Protect(ProtectionType.All);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
