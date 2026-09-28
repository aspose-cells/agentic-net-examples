// Title: Copy a worksheet from one Excel workbook to another while preserving formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy the first sheet of source.xlsx into a new workbook target.xlsx, ensuring all cell formulas are retained. | Create an empty workbook, remove its default sheet, and add a copy of a specific worksheet from another workbook with formulas intact using the Aspose.Cells AddCopy method.
// Common Searches: Aspose.Cells C# copy worksheet to another workbook keep formulas | How to duplicate a sheet with formulas using Aspose.Cells .NET | AddCopy method example preserving formulas in Excel files | Copy first worksheet from source.xlsx to target.xlsx Aspose.Cells
// Tags: Aspose.Cells AddCopy worksheet copy | copy worksheet with formulas .NET | duplicate Excel sheet to new workbook C# | retain formulas during sheet copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the existence of source.xlsx (creating an empty workbook if missing), loads it, creates an empty target workbook, copies the first worksheet using the AddCopy method (which retains all formulas), and saves the result as target.xlsx.
class Program
{
    static void Main()
    {
        // Paths to the source and target Excel files
        string sourcePath = "source.xlsx";
        string targetPath = "target.xlsx";

        try
        {
            // Ensure the source file exists; create an empty workbook if missing
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}. Creating an empty workbook.");
                var emptyWorkbook = new Workbook();
                emptyWorkbook.Worksheets[0].Name = "Sheet1";
                emptyWorkbook.Save(sourcePath);
            }

            // Load the source workbook
            var sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) target workbook and remove the default sheet
            var targetWorkbook = new Workbook();
            targetWorkbook.Worksheets.Clear();

            // Copy the first worksheet from the source workbook to the target workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            targetWorkbook.Worksheets.AddCopy(sourceSheet.Name);

            // Save the target workbook with the copied worksheet
            targetWorkbook.Save(targetPath);
            Console.WriteLine($"Target workbook saved to: {targetPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
