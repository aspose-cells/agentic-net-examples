// Title: How to lock only selected columns in an Excel worksheet using Aspose.Cells for .NET while keeping other cells editable
// AI Prompts: Write C# code with Aspose.Cells that unlocks the entire sheet, locks columns A and C, and then protects the worksheet with a password. | Show how to create a style with IsLocked = true, apply it to specific column ranges, and protect the worksheet using the Aspose.Cells API. | Generate a complete Aspose.Cells example that creates a workbook, fills sample data, locks chosen columns, and saves the file as an .xlsx.
// Common Searches: Aspose.Cells C# lock only column A and C while allowing other columns to be edited | protect worksheet with Aspose.Cells but keep some cells unlocked .NET | how to apply locked style to specific columns using Aspose.Cells in C# | C# Aspose.Cells unlock all cells then lock selected columns before saving
// Tags: Aspose.Cells lock specific columns | worksheet protection with unlocked cells Aspose.Cells | apply locked style to column range C# | selective column locking in Excel using Aspose.Cells | C# Aspose.Cells protect sheet while allowing edits

using Aspose.Cells;
using System;
using System.IO;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, unlocks every cell, applies a locked style to columns A and C, protects the worksheet with a password, and saves the result as LockedColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Salary");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Alice");
            sheet.Cells["C3"].PutValue(6200);

            // ------------------------------------------------------------
            // Step 1: Unlock all cells in the worksheet
            // ------------------------------------------------------------
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // make cells editable
            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true; // apply the Locked property

            // Determine used range (fallback values if sheet is empty)
            int maxRow = sheet.Cells.MaxDataRow >= 0 ? sheet.Cells.MaxDataRow : 1000;
            int maxCol = sheet.Cells.MaxDataColumn >= 0 ? sheet.Cells.MaxDataColumn : 100;

            // Apply the unlocked style to the whole used range
            int totalRows = maxRow + 1; // rows are zero‑based
            int totalCols = maxCol + 1; // columns are zero‑based
            AsposeRange wholeRange = sheet.Cells.CreateRange(0, 0, totalRows, totalCols);
            wholeRange.ApplyStyle(unlockedStyle, unlockedFlag);

            // ------------------------------------------------------------
            // Step 2: Lock specific columns (e.g., Column A and Column C)
            // ------------------------------------------------------------
            int[] columnsToLock = new int[] { 0, 2 }; // 0 = A, 2 = C

            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true; // lock cells
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true; // apply the Locked property

            foreach (int colIndex in columnsToLock)
            {
                // Apply the locked style to the entire column within the used range
                AsposeRange colRange = sheet.Cells.CreateRange(0, colIndex, totalRows, 1);
                colRange.ApplyStyle(lockedStyle, lockedFlag);
            }

            // ------------------------------------------------------------
            // Step 3: Protect the worksheet so that locked cells cannot be edited
            // ------------------------------------------------------------
            sheet.Protect(ProtectionType.All, "StrongPassword123", string.Empty);

            // ------------------------------------------------------------
            // Step 4: Save the workbook
            // ------------------------------------------------------------
            string outputPath = "LockedColumns.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
