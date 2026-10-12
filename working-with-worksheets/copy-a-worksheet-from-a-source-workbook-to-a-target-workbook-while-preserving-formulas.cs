// Title: Copy a worksheet from one Excel workbook to another while preserving formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to duplicate the first sheet of source.xlsx into target.xlsx, keep all formulas, rename the copied sheet, and delete the default Sheet1. | Create a new workbook, add a copy of a specific worksheet from an existing workbook with Worksheets.AddCopy, ensure formulas stay intact, and save the result.
// Common Searches: Aspose.Cells C# copy worksheet to new workbook keep formulas | How to preserve Excel formulas when copying sheets between workbooks with Aspose.Cells | Worksheets.AddCopy example preserving formulas in .NET | Remove default Sheet1 after adding copied worksheet Aspose.Cells
// Tags: Worksheets.AddCopy method C# | copy worksheet with formulas Aspose.Cells | remove default Sheet1 Aspose.Cells | rename copied worksheet Aspose.Cells | duplicate Excel sheet between workbooks .NET

using System;
using System.IO;
using Aspose.Cells;

// The code checks for the source file, loads it into a Workbook, creates an empty target Workbook, copies the first worksheet using Worksheets.AddCopy (renaming it with a _Copy suffix), removes the automatically created Sheet1 if present, and saves the target file, ensuring all formulas are retained.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string targetPath = "target.xlsx";

            // Ensure the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) target workbook
            Workbook targetWorkbook = new Workbook();

            // Get the first worksheet from the source workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Copy the worksheet to the target workbook using the overload that accepts arrays
            string copiedSheetName = sourceSheet.Name + "_Copy";
            targetWorkbook.Worksheets.AddCopy(
                new Worksheet[] { sourceSheet },
                new string[] { copiedSheetName });

            // Remove the default empty sheet that Aspose.Cells creates in a new workbook
            if (targetWorkbook.Worksheets.Count > 1 && targetWorkbook.Worksheets[0].Name == "Sheet1")
            {
                targetWorkbook.Worksheets.RemoveAt(0);
            }

            // Save the target workbook with the copied worksheet
            targetWorkbook.Save(targetPath);
            Console.WriteLine($"Worksheet copied successfully to {targetPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
