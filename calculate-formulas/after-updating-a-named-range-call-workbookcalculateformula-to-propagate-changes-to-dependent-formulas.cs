// Title: Update a named range and recalculate dependent formulas with Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that changes a Name object's RefersTo property to a new range and then invokes Workbook.CalculateFormula to refresh all dependent formulas. | Create a complete Aspose.Cells example that defines a named range, uses it in a formula, updates the range, recalculates the workbook, and saves the file. | Explain the steps required to propagate changes from an updated named range to formulas using Workbook.CalculateFormula in a .NET application.
// Common Searches: Aspose.Cells recalculate formulas after updating a named range in C# | C# Workbook.CalculateFormula after changing Name.RefersTo | How to refresh dependent cells when a named range is modified using Aspose.Cells | Example of updating an Excel named range and recalculating SUM formula with Aspose.Cells .NET
// Tags: modify named range Aspose.Cells C# | Workbook.CalculateFormula usage | refresh dependent formulas after name change | Aspose.Cells update Name.RefersTo | save workbook after recalculation .NET

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeUpdateDemo
{
    // // Example that creates a workbook, defines a named range, uses it in a SUM formula, updates the range reference, calls Workbook.CalculateFormula to recalculate dependent cells, and saves the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data that will be referenced by the named range
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A4"].PutValue(40);
            cells["A5"].PutValue(50);

            // Add a named range that initially refers to A1:A3
            int nameIndex = wb.Worksheets.Names.Add("MyRange");
            Name myRange = wb.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula (sum of the range)
            cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate formulas before the named range change
            wb.CalculateFormula(); // lifecycle rule: calculate
            Console.WriteLine("Sum before range update: " + cells["B1"].Value); // Expected 60 (10+20+30)

            // Update the named range to refer to A1:A5
            myRange.RefersTo = "=Sheet1!$A$1:$A$5";

            // Propagate the change by recalculating formulas (required after named range update)
            wb.CalculateFormula(); // lifecycle rule: calculate

            // Output the new result
            Console.WriteLine("Sum after range update: " + cells["B1"].Value); // Expected 150 (10+20+30+40+50)

            // Save the workbook (lifecycle rule: save)
            wb.Save("NamedRangeUpdateResult.xlsx");
        }
    }
}
