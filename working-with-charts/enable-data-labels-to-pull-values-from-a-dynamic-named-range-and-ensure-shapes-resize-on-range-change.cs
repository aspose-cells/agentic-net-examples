// Title: Aspose.Cells .NET: Bind Chart Data Labels to a Dynamic Named Range & Auto‑Resize Shapes
// Description: Demonstrates how to create a column chart, link its data labels to a FILTER‑based dynamic named range, enable each label shape to resize and auto‑scale its font, refresh spilled array formulas, and recalculate the chart after the source data changes, saving both the initial and updated workbooks.
// Keywords: Aspose.Cells chart data labels | dynamic named range .NET | FILTER formula spill range | auto‑resize label shapes | RefreshDynamicArrayFormulas | Chart.Calculate Aspose.Cells | C# Aspose.Cells example
// Common Searches: Aspose.Cells bind chart labels to dynamic range | auto resize data label shapes Aspose.Cells | refresh spilled array formulas .NET | link chart data labels to FILTER result | update chart after adding rows Aspose.Cells
// Developer Intent: Create a column chart whose labels are driven by a dynamic named range and automatically adjust their shape size when the range expands or contracts.
// Use Cases: Link chart data labels to a FILTER‑based dynamic array so they update with added or removed rows. | Set IsResizeShapeToFitText and AutoScaleFont on each ChartPoint to keep label shapes fitting their content. | Refresh dynamic array formulas and recalculate the chart after modifying source data to reflect new label values.
// AI Prompts: Generate C# code with Aspose.Cells that creates a chart whose data labels are linked to a dynamic named range and resize automatically. | Show how to enable IsResizeShapeToFitText and AutoScaleFont for ChartPoint data labels in Aspose.Cells .NET. | Explain the steps to refresh spilled array formulas and recalculate a chart after adding rows to the worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDynamicDataLabels
{
    // Demonstrates how to create a column chart, link its data labels to a FILTER‑based dynamic named range, enable each label shape to resize and auto‑scale its font, refresh spilled array formulas, and recalculate the chart after the source data changes, saving both the initial and updated workbooks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (categories in A, values in B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10); // 20,30,40,50,60
            }

            // Set a dynamic array formula in D1 that spills the values from column B
            // This creates a dynamic range that expands/shrinks when column B changes
            sheet.Cells["D1"].SetDynamicArrayFormula("=FILTER(B2:B6, A2:A6<>\"\")", new FormulaParseOptions(), true);

            // Define a named range that points to the spilled range (using the # operator)
            int nameIndex = workbook.Worksheets.Names.Add("LabelValues");
            Name labelName = workbook.Worksheets.Names[nameIndex];
            labelName.RefersTo = "=Sheet1!$D$1#";

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Set chart data source
            chart.NSeries.Add("B2:B6", true);          // Values
            chart.NSeries.CategoryData = "A2:A6";      // Categories

            // Configure data labels to pull values from the dynamic named range
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = false;               // Hide default values
            series.DataLabels.ShowCellRange = true;            // Enable showing cell range
            series.DataLabels.LinkedSource = "LabelValues";    // Link to dynamic named range

            // Ensure each data label shape resizes to fit the text when the range changes
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsResizeShapeToFitText = true;
                // Optional: let the shape auto-size (width/height will adapt)
                point.DataLabels.AutoScaleFont = true;
            }

            // Calculate the chart to apply the data label texts
            chart.Calculate();

            // Save the initial workbook
            workbook.Save("DynamicDataLabels_Initial.xlsx");

            // ----- Simulate a change in source data that expands the dynamic range -----
            // Add more rows of data
            for (int i = 7; i <= 9; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10); // 70,80,90
            }

            // Refresh dynamic array formulas so the spill range updates
            workbook.RefreshDynamicArrayFormulas(true);
            // Recalculate formulas (optional but ensures all dependent formulas are up‑to‑date)
            workbook.CalculateFormula();

            // Re‑calculate the chart so data labels reflect the new spilled values
            chart.Calculate();

            // Save the workbook after the data change
            workbook.Save("DynamicDataLabels_Updated.xlsx");
        }
    }
}
