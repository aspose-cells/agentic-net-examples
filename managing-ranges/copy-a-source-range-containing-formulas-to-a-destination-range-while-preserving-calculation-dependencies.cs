// Title: Copy a Range with Formulas and Preserve Dependencies – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a source range that contains formulas to another location using Aspose.Cells' `Range.Copy` method, keep relative references intact, recalculate the workbook, and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | Range.Copy | copy formulas | preserve dependencies | recalculate formulas | Excel automation | XLSX export | relative cell references
// Common Searches: Aspose.Cells copy range with formulas C# | how to keep formula references when copying cells Aspose.Cells | Range.Copy preserve dependencies .NET | recalculate workbook after copying formulas Aspose | copy A1:B5 to A7:B11 Aspose.Cells example
// Developer Intent: Copy a block of cells that includes formulas to a new area while maintaining correct relative references and updating calculation results.
// Use Cases: Duplicate a calculation block for scenario analysis without breaking formula links. | Create multiple template sections for different data sets while keeping formulas functional. | Generate a summary table by copying a formula range to another sheet and refreshing values.
// AI Prompts: Provide C# code that uses Aspose.Cells to copy a range with formulas and automatically adjusts references. | Show how to recalculate all formulas after copying a range with `Range.Copy` in Aspose.Cells. | Explain the handling of relative and absolute references when a range containing formulas is copied with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to copy a source range that contains formulas to another location using Aspose.Cells' `Range.Copy` method, keep relative references intact, recalculate the workbook, and save the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source range A1:A5 with numeric values
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // A1..A5 = 1,2,3,4,5
            }

            // Add formulas in B1:B5 that depend on the values in column A
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 1].Formula = $"A{i + 1}*2"; // B column = A*2
            }

            // Define source range (A1:B5) and destination range (A7:B11)
            AsposeRange sourceRange = sheet.Cells.CreateRange(0, 0, 5, 2);   // rows 0-4, cols 0-1
            AsposeRange destinationRange = sheet.Cells.CreateRange(6, 0, 5, 2); // rows 6-10, cols 0-1

            // Copy the source range to the destination range, preserving formulas and dependencies
            sourceRange.Copy(destinationRange);

            // Recalculate formulas so that the copied range shows correct results
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("CopyRangeWithFormulas.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
