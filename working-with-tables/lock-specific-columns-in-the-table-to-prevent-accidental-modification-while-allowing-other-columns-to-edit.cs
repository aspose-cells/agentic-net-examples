// Title: Lock selected columns in an Excel worksheet while allowing other columns to be edited with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock column A and keep other columns editable | protect Excel worksheet with password but unlock specific columns using Aspose.Cells | apply locked style to particular columns in an Aspose.Cells workbook C# example | make only certain columns read‑only in an Excel file with Aspose.Cells .NET | C# code to lock columns and protect worksheet using Aspose.Cells
// Tags: lock columns Aspose.Cells .NET | worksheet protection password Aspose.Cells | apply style flag locked Aspose.Cells | unlock all cells Aspose.Cells | column‑level read‑only Excel Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads or creates a workbook, unlocks every cell, applies a locked style to the specified columns (e.g., A and C), protects the worksheet with a password, and saves the file, ensuring only those columns are read‑only.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load an existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Columns to lock (0‑based index). Example: lock Column A (0) and Column C (2)
            int[] columnsToLock = new int[] { 0, 2 };

            // -----------------------------------------------------------------
            // Step 1: Unlock all cells first so that only selected columns become locked
            // -----------------------------------------------------------------
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;               // Cells become editable
            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true;                    // Apply the Locked property

            // Apply the unlocked style to the entire sheet
            worksheet.Cells.ApplyStyle(unlockedStyle, unlockedFlag);

            // -----------------------------------------------------------------
            // Step 2: Lock the specified columns
            // -----------------------------------------------------------------
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true;                  // Cells become locked
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true;                     // Apply the Locked property

            foreach (int colIndex in columnsToLock)
            {
                // Apply the locked style to the whole column
                worksheet.Cells.Columns[colIndex].ApplyStyle(lockedStyle, lockedFlag);
            }

            // -----------------------------------------------------------------
            // Step 3: Protect the worksheet to enforce the lock settings
            // -----------------------------------------------------------------
            // Provide a password; oldPassword is not required for new protection, pass empty string
            worksheet.Protect(ProtectionType.All, "myPassword", string.Empty);

            // -----------------------------------------------------------------
            // Step 4: Save the modified workbook
            // -----------------------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
