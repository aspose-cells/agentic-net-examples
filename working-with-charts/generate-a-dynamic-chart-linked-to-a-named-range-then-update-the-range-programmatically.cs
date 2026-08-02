// Title: C# – Create a Dynamic Column Chart Linked to a Named Spill Range and Update It with Aspose.Cells
// Description: Demonstrates how to build a workbook, set a SEQUENCE dynamic array in B2, define a named range that points to the spill range (using the # operator), attach a column chart to that named range, save the file, then programmatically change the row count, refresh dynamic array formulas, recalculate, and save the updated workbook so the chart reflects the new data.
// Keywords: Aspose.Cells C# dynamic chart | named range spill range | SEQUENCE function chart data | refresh dynamic array formulas | programmatic chart update | column chart from dynamic array | Aspose.Cells chart binding
// Common Searches: bind Aspose.Cells chart to spill range using named range | expand SEQUENCE array and refresh chart in .NET | dynamic chart data source Aspose.Cells C# | update chart after changing source cell Aspose.Cells | create column chart from dynamic array formula
// Developer Intent: Generate a column chart that reads from a dynamic spill range via a named range and automatically reflects changes made to the source array.
// Use Cases: Automatically grow a chart when the number of rows returned by a SEQUENCE formula changes. | Use a named range with the # spill operator to keep chart data linked to a dynamic array. | Refresh dynamic array formulas and recalculate the workbook to update chart visuals without manual steps.
// AI Prompts: Show how to bind an Aspose.Cells column chart to a named range that references a spill range (#) in C#. | Provide code to change the size of a SEQUENCE‑based dynamic array and refresh the linked chart using Aspose.Cells. | Explain the steps required for a chart to update automatically after modifying the cell that controls a dynamic array.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartDemo
{
    // Demonstrates how to build a workbook, set a SEQUENCE dynamic array in B2, define a named range that points to the spill range (using the # operator), attach a column chart to that named range, save the file, then programmatically change the row count, refresh dynamic array formulas, recalculate, and save the updated workbook so the chart reflects the new data.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Prepare data for dynamic array ----------
            // Cell B1 will hold the number of rows for the SEQUENCE function
            cells["B1"].PutValue(5);                     // initial count = 5
            // Set a dynamic array formula in B2 that spills vertically
            cells["B2"].SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // ---------- Create a named range that points to the spill range ----------
            // The spill range is referenced with the # symbol
            int nameIdx = wb.Worksheets.Names.Add("ChartData");
            wb.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!$B$2#";

            // ---------- Add a column chart that uses the named range ----------
            // Use the Add method that accepts dataRange (named range) and orientation flag
            int chartIdx = sheet.Charts.Add(ChartType.Column, "ChartData", true, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Dynamic Data Chart";

            // ---------- Save the initial workbook ----------
            wb.Save("DynamicChart_Initial.xlsx");

            // ---------- Update the dynamic array range programmatically ----------
            // Change the count in B1 to expand the spill range
            cells["B1"].PutValue(8); // new count = 8

            // Refresh dynamic array formulas so the spill range updates
            wb.RefreshDynamicArrayFormulas(true);

            // Optionally recalculate the workbook (not strictly required for the chart)
            wb.CalculateFormula();

            // ---------- Save the workbook after update ----------
            wb.Save("DynamicChart_Updated.xlsx");
        }
    }
}
