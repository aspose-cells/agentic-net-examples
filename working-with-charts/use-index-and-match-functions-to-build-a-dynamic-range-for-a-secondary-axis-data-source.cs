// Title: C# – Use INDEX + MATCH to Define a Dynamic Secondary Axis Range in an Aspose.Cells Column Chart
// Description: This example shows how to create a workbook, populate primary and secondary data, add a column chart, enable a secondary axis, and assign the secondary category series to a cell reference generated at runtime with an INDEX/MATCH formula. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | C# chart secondary axis | dynamic range INDEX MATCH | column chart formula | secondary category data | Excel chart automation .NET | runtime chart data source
// Common Searches: Aspose.Cells set secondary axis range with formula | C# INDEX MATCH chart data source Aspose | dynamic secondary category in Excel chart .NET | how to bind secondary axis to a formula in Aspose.Cells | create column chart with dynamic secondary axis C#
// Developer Intent: Assign a formula‑driven, runtime‑calculated range to the secondary axis of an Aspose.Cells chart.
// Use Cases: Build a sales chart where the secondary axis labels change automatically based on a month selected in a cell. | Generate regional reports that switch the displayed region on the secondary axis without code changes. | Create dashboards that adapt their secondary series labels when users modify a lookup value.
// AI Prompts: Generate C# code that sets a chart's secondary category data to an INDEX/MATCH formula using Aspose.Cells. | Explain how to extend the formula to return a multi‑cell range for the secondary axis. | Provide steps to refresh the secondary axis automatically when the lookup cell value is updated at runtime.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicSecondaryAxis
{
    // This example shows how to create a workbook, populate primary and secondary data, add a column chart, enable a secondary axis, and assign the secondary category series to a cell reference generated at runtime with an INDEX/MATCH formula. The workbook is then saved as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Populate sample data
            // ------------------------------------------------------------
            // Primary categories (A column)
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["A5"].PutValue("Apr");

            // Primary series values (B column)
            cells["B1"].PutValue("Sales");
            cells["B2"].PutValue(120);
            cells["B3"].PutValue(150);
            cells["B4"].PutValue(130);
            cells["B5"].PutValue(170);

            // Secondary categories (C column) – we will select a subset dynamically
            cells["C1"].PutValue("Region");
            cells["C2"].PutValue("North");
            cells["C3"].PutValue("South");
            cells["C4"].PutValue("East");
            cells["C5"].PutValue("West");

            // Lookup value that determines which secondary category to start from (D1)
            cells["D1"].PutValue("Feb"); // Change this value to test dynamic range

            // ------------------------------------------------------------
            // Add a column chart
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add primary series (sales values)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable secondary axis for the first series
            chart.NSeries[0].PlotOnSecondAxis = true;

            // ------------------------------------------------------------
            // Build a dynamic range for secondary axis data using INDEX+MATCH
            // ------------------------------------------------------------
            // Formula explanation:
            //   MATCH(D1, A2:A5, 0)  -> finds the row index of the lookup value in column A
            //   INDEX(C2:C5, MATCH(...)) -> returns the corresponding secondary category
            // This example sets the secondary category data to a single cell determined at runtime.
            // For a true range you could combine two INDEX calls, but Aspose.Cells accepts a single-cell reference here.
            string dynamicRangeFormula = "=INDEX($C$2:$C$5, MATCH($D$1, $A$2:$A$5, 0))";

            // Assign the dynamic range to the secondary category axis
            chart.NSeries.SecondCategoryData = dynamicRangeFormula;

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("DynamicSecondaryAxis.xlsx");
        }
    }
}
