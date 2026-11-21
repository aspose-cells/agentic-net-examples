// Title: Freeze the first row and protect only the header in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze top row and lock header while keeping other cells editable | How to protect worksheet header after using FreezePanes in Aspose.Cells .NET | C# code to unlock all cells then lock the first row with Aspose.Cells | Apply worksheet protection with locked header after freezing panes using Aspose.Cells | Aspose.Cells protect sheet but allow editing of data rows C#
// Tags: freeze panes protect header Aspose.Cells | lock header row worksheet Aspose.Cells | unlock all cells then lock first row Aspose.Cells | apply style to range Aspose.Cells C# | worksheet protection all actions Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, freezes the first row, unlocks every cell, locks only the header row, applies worksheet protection, and saves the file as ProtectedSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Freeze the top row (header)
            sheet.FreezePanes(0, 0, 1, 0);

            // Sample data (optional, just to illustrate)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["A3"].PutValue("Row 2");

            // ---------- Unlock all cells ----------
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // make cells editable
            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true; // apply the Locked property

            // Apply the unlocked style to the whole used range
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalCols = sheet.Cells.MaxDataColumn + 1;
            AsposeRange allRange = sheet.Cells.CreateRange(0, 0, totalRows, totalCols);
            allRange.ApplyStyle(unlockedStyle, unlockedFlag);

            // ---------- Lock the header row ----------
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true; // lock header cells
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true; // apply the Locked property

            // Apply the locked style to the first row
            AsposeRange headerRange = sheet.Cells.CreateRange(0, 0, 1, totalCols);
            headerRange.ApplyStyle(lockedStyle, lockedFlag);

            // Protect the worksheet so locked cells cannot be edited
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "ProtectedSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
