// Title: Aspose.Cells .NET – Column chart with data‑label cells from a dynamic range and auto‑sized label shapes
// Description: C# code that builds a workbook, defines a dynamic named range via OFFSET/COUNTA, adds a column chart linked to that range, sets data labels to read directly from the cells, preserves number formatting, enables automatic shape resizing, and refreshes formulas and the chart after new rows are added.
// Keywords: Aspose.Cells | .NET chart automation | dynamic named range | OFFSET function | COUNTA function | linked data labels | auto‑resize label shape | RefreshDynamicArrayFormulas | chart.Calculate | column chart
// Common Searches: Aspose.Cells bind chart series to a dynamic named range | how to link chart data labels to cells in Aspose.Cells | auto‑size data‑label shapes after range expansion .NET | refresh chart after adding rows with Aspose.Cells | dynamic range chart example C# Aspose
// Developer Intent: Generate a column chart whose labels are sourced from a growing named range and automatically adjust their shape size as the range changes.
// Use Cases: Create reports where new numeric entries are appended and chart labels must reflect the latest values without manual updates. | Maintain consistent number formatting on labels while the underlying data range expands. | Automate workbook recalculation and chart refresh after inserting additional rows.
// AI Prompts: Write C# using Aspose.Cells to build a column chart with data labels linked to a dynamic OFFSET named range and enable auto‑fit of label shapes. | Explain how to refresh dynamic array formulas and recalculate a chart after extending the source range in Aspose.Cells for .NET. | Provide step‑by‑step code to bind a chart series to a named range defined by OFFSET/COUNTA and activate linked data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicDataLabels
{
    // C# code that builds a workbook, defines a dynamic named range via OFFSET/COUNTA, adds a column chart linked to that range, sets data labels to read directly from the cells, preserves number formatting, enables automatic shape resizing, and refreshes formulas and the chart after new rows are added.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column B (values) and column C (formatted text)
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue("10 units");
            sheet.Cells["C3"].PutValue("20 units");
            sheet.Cells["C4"].PutValue("30 units");

            // Define a dynamic named range "DynValues" that expands based on non‑empty cells in column B
            // =OFFSET(Sheet1!$B$2,0,0,COUNTA(Sheet1!$B$2:$B$100),1)
            int nameIndex = workbook.Worksheets.Names.Add("DynValues");
            Name dynName = workbook.Worksheets.Names[nameIndex];
            dynName.RefersTo = "=OFFSET(Sheet1!$B$2,0,0,COUNTA(Sheet1!$B$2:$B$100),1)";

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Use the dynamic named range for the series values
            chart.NSeries.Add("DynValues", true);
            // Category data (optional, using static range for simplicity)
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to pull values from a dynamic linked source
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.ShowCellRange = true;           // Enable pulling from a cell range
            series.DataLabels.LinkedSource = "DynValues";     // Link to the dynamic named range
            series.DataLabels.NumberFormatLinked = true;      // Keep number format in sync
            series.DataLabels.IsResizeShapeToFitText = true;  // Auto‑fit shape to label text

            // Initial calculation to render the chart correctly
            chart.Calculate();

            // Simulate a change that expands the data range
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["C5"].PutValue("40 units");

            // Refresh dynamic array formulas (if any) and recalculate the workbook
            workbook.RefreshDynamicArrayFormulas(true);
            workbook.CalculateFormula();

            // Re‑calculate the chart so that data labels reflect the updated range
            chart.Calculate();

            // Save the workbook
            workbook.Save("DynamicDataLabels.xlsx");
        }
    }
}
