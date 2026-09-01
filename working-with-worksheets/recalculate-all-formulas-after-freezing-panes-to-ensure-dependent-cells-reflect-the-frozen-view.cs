// Title: Recalculate all formulas after freezing panes with Aspose.Cells in C#
// AI Prompts: Generate C# code that applies Worksheet.FreezePanes and then forces a full workbook calculation using Aspose.Cells. | Show how to recalculate dependent formulas after freezing rows and columns in an Aspose.Cells workbook. | Provide a complete example that saves the workbook after the second CalculateFormula call in Aspose.Cells for .NET. | Explain why a second CalculateFormula call is required after FreezePanes and demonstrate it in C#.
// Common Searches: Aspose.Cells C# recalculate formulas after using FreezePanes | How to update dependent cells when freezing panes in a .NET Excel workbook | Worksheet.FreezePanes followed by CalculateFormula example Aspose.Cells | C# Aspose.Cells refresh formulas after setting frozen rows and columns | Recalculate all formulas after applying FreezePanes in Aspose.Cells
// Tags: Worksheet.FreezePanes formula recalculation | Aspose.Cells CalculateFormula after freeze panes | C# Excel freeze panes dependent cell update | Aspose.Cells workbook recalc post freeze | Save workbook after recalculation Aspose.Cells

using System;
using Aspose.Cells;

namespace RecalculateAfterFreezePanes
{
    // The program creates a new workbook, fills cells A1‑A3 with numeric values, adds formulas in B1‑B3 and a SUM in C1, performs an initial calculation, freezes the top row and first two columns at cell C2 using Worksheet.FreezePanes, calls CalculateFormula again to ensure all dependent formulas reflect the frozen view, prints key cell values to the console, and saves the file as RecalculatedAfterFreezePanes.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add formulas that depend on the data above
            cells["B1"].Formula = "=A1*2";   // 20
            cells["B2"].Formula = "=A2*2";   // 40
            cells["B3"].Formula = "=A3*2";   // 60
            cells["C1"].Formula = "=SUM(B1:B3)"; // 120

            // Initial calculation to establish values and dependency chain
            workbook.CalculateFormula();

            // Freeze panes at cell C2 (row index 1, column index 2) with 1 frozen row and 2 frozen columns
            // This ensures the top row and left two columns stay visible while scrolling
            sheet.FreezePanes(1, 2, 1, 2);

            // Recalculate all formulas after freezing panes
            // This guarantees that any dependent cells reflect the current view state
            workbook.CalculateFormula();

            // Output results to console for verification
            Console.WriteLine("A1: " + cells["A1"].Value);
            Console.WriteLine("B1: " + cells["B1"].Value);
            Console.WriteLine("C1 (SUM): " + cells["C1"].Value);

            // Save the workbook
            workbook.Save("RecalculatedAfterFreezePanes.xlsx");
        }
    }
}
