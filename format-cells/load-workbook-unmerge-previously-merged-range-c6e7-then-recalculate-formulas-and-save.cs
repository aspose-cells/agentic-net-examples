// Title: How to unmerge the cell range C6:E7, recalculate all formulas, and save the workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to unmerge cells C6 to E7 with Aspose.Cells and recalculate workbook formulas | Aspose.Cells .NET example for removing a merged range and updating formulas before saving | Programmatically split a specific Excel range and force formula calculation using Aspose.Cells in C# | How to use Aspose.Cells to unmerge cells C6:E7 and recalc all formulas | Save an Excel file after unmerging cells and recalculating formulas with Aspose.Cells for .NET
// Tags: Aspose.Cells unmerge range C# | Aspose.Cells calculate formulas .NET | Aspose.Cells create range and unmerge | Aspose.Cells save workbook after unmerge | Aspose.Cells modify merged cells programmatically

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook (creating a blank file if missing), accesses the first worksheet, defines the range C6:E7, unmerges that range, recalculates all workbook formulas, and saves the updated file to a new location.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Save(inputPath);
            }

            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the previously merged range C6:E7.
            // startRow = 5 (row 6), startColumn = 2 (column C)
            // totalRows = 2 (rows 6‑7), totalColumns = 3 (columns C‑E)
            var range = sheet.Cells.CreateRange(5, 2, 2, 3);
            range.UnMerge();

            // Recalculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
