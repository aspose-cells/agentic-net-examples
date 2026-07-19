// Title: Create a dynamic column chart linked to a named range using Aspose.Cells for .NET
// Description: This example shows how to build a workbook, define a named range (MyData), spill the range with SetDynamicArrayFormula, add a column chart, bind the chart to the named range via SetChartDataRange, modify the range values in code, refresh dynamic array formulas, recalculate the chart, and save the file.
// Keywords: Aspose.Cells dynamic chart | named range chart .NET | SetDynamicArrayFormula C# | SetChartDataRange example | refresh dynamic array formulas | programmatic chart update | Aspose.Cells C# tutorial | Excel chart linked to named range
// Common Searches: Aspose.Cells bind chart to named range | How to refresh dynamic array formulas in Aspose.Cells | Update chart data programmatically in C# | Create column chart from named range Aspose.Cells | Dynamic chart example Aspose.Cells .NET
// Developer Intent: Generate a column chart that reads data from a named range, change the range values in code, and refresh the chart so the visual reflects the new data.
// Use Cases: Automated sales dashboards where chart data updates automatically after a nightly data load. | Financial reports that use a named range to centralize figures, keeping charts in sync without manual re‑linking. | Interactive Excel workbooks that let users edit source cells while the chart redraws instantly via code.
// AI Prompts: Write C# code with Aspose.Cells to create a named range, spill it using SetDynamicArrayFormula, attach a column chart, modify the range values, refresh formulas, and recalculate the chart. | Explain the relationship between SetChartDataRange and RefreshDynamicArrayFormulas for keeping a chart current when the underlying named range changes. | Provide a step‑by‑step guide to bind a dynamic array‑spilled named range to a chart, then programmatically update the data and refresh the visualization.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartDemo
{
    // This example shows how to build a workbook, define a named range (MyData), spill the range with SetDynamicArrayFormula, add a column chart, bind the chart to the named range via SetChartDataRange, modify the range values in code, refresh dynamic array formulas, recalculate the chart, and save the file.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data that will be referenced by a named range
            //    A2:A5 – categories, B2:B5 – values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // 3. Define a named range that points to the data area (A2:B5)
            int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyData");
            sheet.Workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$2:$B$5";

            // 4. Set a dynamic array formula that simply returns the named range.
            //    The formula will spill into the cells starting at E2.
            //    Using SetDynamicArrayFormula (rule) with calculateValue = true.
            Cell dynamicCell = sheet.Cells["E2"];
            dynamicCell.SetDynamicArrayFormula("=MyData", new FormulaParseOptions(), true);

            // 5. Add a column chart to the worksheet (rule: ChartCollection.Add)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // 6. Link the chart to the named range using SetChartDataRange (rule)
            //    The range string can be the name itself.
            chart.SetChartDataRange("MyData", true);

            // 7. Update the underlying data programmatically.
            //    Change the values in the named range to demonstrate dynamic update.
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(35);
            sheet.Cells["B5"].PutValue(45);

            // 8. Refresh dynamic array formulas so the spilled range reflects the new data.
            workbook.RefreshDynamicArrayFormulas(true);

            // 9. Recalculate the chart to ensure it picks up the modified data.
            chart.Calculate();

            // 10. Save the workbook (standard save operation)
            workbook.Save("DynamicChartLinkedToNamedRange.xlsx");
        }
    }
}
