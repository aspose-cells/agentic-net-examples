// Title: Define a global named range across multiple worksheets and sum it on a summary sheet using Aspose.Cells for .NET
// AI Prompts: Create a named range that includes A1:A2 on Sheet1 and Sheet2, then insert a =SUM(MyMultiSheetRange) formula in cell A1 of a new Summary worksheet and evaluate the workbook. | Add a named range that spans several sheets, call workbook.CalculateFormula(), and save the workbook as an .xlsx file using Aspose.Cells in C#.
// Common Searches: Aspose.Cells how to create a named range that spans more than one worksheet in C# | C# sum values from cells on different sheets using a named range with Aspose.Cells | global named range across multiple sheets Aspose.Cells .NET example | calculate formulas after defining a multi‑sheet named range in Aspose.Cells | save workbook after using SUM formula with a multi‑sheet named range in C#
// Tags: Aspose.Cells create multi‑sheet named range | global scope named range .NET Aspose.Cells | use SUM with named range Aspose.Cells | C# workbook formula calculation Aspose.Cells | export workbook to .xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates creating a workbook with two worksheets, defining a global named range that covers A1:A2 on both sheets, adding a Summary sheet that uses =SUM(MyMultiSheetRange) to total the values, calculating the formula, and saving the file as NamedRangeMultiSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate sample data in both worksheets
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet2.Cells["A1"].PutValue(30);
            sheet2.Cells["A2"].PutValue(40);

            // Create a named range that spans A1:A2 on both sheets
            string rangeReference = "Sheet1!A1:A2,Sheet2!A1:A2";

            // Add the named range (global scope) and set its reference
            int nameIndex = workbook.Worksheets.Names.Add("MyMultiSheetRange");
            Name multiSheetRange = workbook.Worksheets.Names[nameIndex];
            multiSheetRange.RefersTo = rangeReference;

            // Add a summary worksheet
            Worksheet summary = workbook.Worksheets.Add("Summary");

            // Use the named range in a formula to sum all values across the sheets
            summary.Cells["A1"].Formula = "=SUM(MyMultiSheetRange)";

            // Calculate formulas so the result is stored in the cell
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "NamedRangeMultiSheet.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
