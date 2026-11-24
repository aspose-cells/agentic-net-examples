// Title: Programmatically lock cells containing formulas and protect worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# lock cells that have formulas and protect the worksheet | how to set IsLocked style for formula cells with Aspose.Cells .NET | protect Excel sheet after locking formula cells using Aspose.Cells API | iterate used range to lock formula cells in a workbook with Aspose.Cells | C# code to prevent editing of formula cells in an existing XLSX file
// Tags: lock formula cells Aspose.Cells | worksheet protection Aspose.Cells .NET | set cell IsLocked style C# | detect formulas in used range Aspose.Cells | save protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX workbook, iterates each worksheet's used range, identifies cells with formulas, sets their Style.IsLocked flag to true, applies full worksheet protection, and saves the protected file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // If the cell contains a formula, lock it
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Style style = cell.GetStyle();
                        style.IsLocked = true;
                        cell.SetStyle(style);
                    }
                }

                // Protect the worksheet to enforce the locked cells
                sheet.Protect(ProtectionType.All);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
