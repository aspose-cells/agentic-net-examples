// Title: C# Aspose.Cells: Verify Chart Data Labels Update After Changing Linked Cells
// Description: Creates a column chart, links its data labels to cells C2:C3, updates those cells, calls chart.Calculate(), and iterates through each point to confirm the displayed label matches the new cell value, saving workbooks before and after.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# chart data labels | linked cell range | chart.Calculate() | validate data label text | column chart | programmatic cell update | Excel chart automation | Aspose.Cells API
// Common Searches: Aspose.Cells chart data labels linked cells C# | how to refresh chart after cell change Aspose.Cells | verify data label values programmatically Aspose.Cells | C# example validate chart labels after updating cells | Aspose.Cells recalculate chart for updated labels
// Developer Intent: Confirm that chart data labels reflect the latest values from their linked cells after the cells are modified via code.
// Use Cases: Link data labels to a worksheet range and programmatically change the range values. | Refresh the chart with chart.Calculate() to apply cell updates. | Iterate through series points and compare DataLabels.Text with the corresponding cell content to detect mismatches. | Generate before‑and‑after workbooks to audit label changes.
// AI Prompts: Provide C# code using Aspose.Cells to link chart data labels to a cell range, modify the cells, call chart.Calculate(), and assert that each label matches the updated cell value. | Create a function that accepts a Chart object and a linked cell address, updates the cells, recalculates the chart, and returns true if all data labels are synchronized. | Explain the process of validating Aspose.Cells chart data labels after programmatic cell updates, including required properties and methods.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelValidation
{
    // Creates a column chart, links its data labels to cells C2:C3, updates those cells, calls chart.Calculate(), and iterates through each point to confirm the displayed label matches the new cell value, saving workbooks before and after.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);

            // Cells that will be linked to data labels
            worksheet.Cells["C2"].PutValue("100 units");
            worksheet.Cells["C3"].PutValue("200 units");

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Configure data labels to show linked cell values
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.ShowCellRange = true;           // Enable showing linked cell range
            series.DataLabels.LinkedSource = "C2:C3";         // Link to cells C2:C3
            series.DataLabels.Font.Color = Color.Blue;       // Optional styling

            // Save the workbook before modification (optional)
            workbook.Save("DataLabelsBeforeModification.xlsx");

            // ---- Modify the linked source cells programmatically ----
            worksheet.Cells["C2"].PutValue("150 units");
            worksheet.Cells["C3"].PutValue("250 units");

            // Recalculate the chart to reflect changes
            chart.Calculate();

            // Verify that each data label now displays the updated cell values
            bool allLabelsCorrect = true;
            for (int i = 0; i < series.Points.Count; i++)
            {
                // Retrieve the displayed text of the data label for the point
                string labelText = series.Points[i].DataLabels.Text;

                // Expected text comes from the linked source cells
                string expected = worksheet.Cells[i + 2, 2].StringValue; // C2, C3 ...

                if (labelText != expected)
                {
                    allLabelsCorrect = false;
                    Console.WriteLine($"Mismatch at point {i}: label='{labelText}' expected='{expected}'");
                }
                else
                {
                    Console.WriteLine($"Point {i} label correctly shows '{labelText}'");
                }
            }

            Console.WriteLine(allLabelsCorrect
                ? "All data labels display the correct updated cell values."
                : "Some data labels do not match the updated cell values.");

            // Save the final workbook
            workbook.Save("DataLabelsAfterModification.xlsx");
        }
    }
}
