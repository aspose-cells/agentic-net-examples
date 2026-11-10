// Title: How to lock all cells except a named range in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# lock entire worksheet and keep a specific range editable | unlock cells in range A1:B2 after protecting sheet with Aspose.Cells | apply locked style to all cells then exclude a range using Aspose.Cells .NET | protect Excel sheet with password while allowing edits in a defined range Aspose.Cells | programmatically set cell lock property for whole sheet except a named range in C#
// Tags: apply locked style to entire worksheet Aspose.Cells | unlock cells within defined range Aspose.Cells | protect worksheet with password Aspose.Cells | range-based cell lock configuration Aspose.Cells | create range and modify style Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, applies a locked style to every cell, creates the range A1:B2, unlocks cells inside that range, protects the worksheet with a password, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Ensure every cell on the sheet is locked (default style)
            // -----------------------------------------------------------------
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true; // cells are locked

            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true;    // apply the lock flag
            sheet.Cells.ApplyStyle(lockedStyle, lockedFlag);

            // -----------------------------------------------------------------
            // 2. Define the range that should stay editable
            // -----------------------------------------------------------------
            const string rangeAddress = "A1:B2"; // relative to the sheet

            // -----------------------------------------------------------------
            // 3. Unlock cells inside the defined range
            // -----------------------------------------------------------------
            Aspose.Cells.Range editableRange = sheet.Cells.CreateRange(rangeAddress);

            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // cells are unlocked

            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true;     // apply the lock flag
            editableRange.ApplyStyle(unlockedStyle, unlockedFlag);

            // -----------------------------------------------------------------
            // 4. Protect the worksheet so that locked cells cannot be edited
            // -----------------------------------------------------------------
            // oldPassword is empty, new password is "password"
            sheet.Protect(ProtectionType.All, "", "password");

            // -----------------------------------------------------------------
            // 5. Save the modified workbook
            // -----------------------------------------------------------------
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
