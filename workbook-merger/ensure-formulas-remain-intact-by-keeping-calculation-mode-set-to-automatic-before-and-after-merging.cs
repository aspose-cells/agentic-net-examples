// Title: Preserve automatic calculation mode while merging a cell range in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an existing workbook, stores the current Workbook.Settings.CalcMode, sets it to Automatic, merges cells A1:C3 on the first worksheet, then restores the original CalcMode before saving with Aspose.Cells. | Show a method that temporarily forces automatic calculation during any cell merge operation and reverts to the previous calculation mode after the merge using the Aspose.Cells API. | Create a reusable helper that ensures formulas stay intact by preserving the workbook's calculation mode when performing a Cells.Merge call in C#.
// Common Searches: Aspose.Cells C# keep calculation mode automatic when merging cells | how to preserve formulas after merging cells with Aspose.Cells .NET | set Workbook.Settings.CalcMode before Cells.Merge and restore after | prevent formula loss during cell range merge in Aspose.Cells | C# Aspose.Cells merge A1:C3 without disabling automatic calculations
// Tags: automatic calculation mode Aspose.Cells | cell merge keep formulas | Workbook.Settings.CalcMode handling | Aspose.Cells merge cells without disabling calculations | C# preserve formula integrity during merge

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), temporarily switches the workbook’s calculation mode to Automatic, merges cells A1:C3 on the first worksheet, restores the original calculation mode, and saves the result, ensuring any formulas remain intact.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Perform merge on the first worksheet (A1:C3)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.Merge(0, 0, 3, 3); // startRow, startColumn, totalRows, totalColumns

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
