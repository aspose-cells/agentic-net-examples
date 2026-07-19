// Title: Create a Dynamic Excel Column Chart from an In‑Memory List with Aspose.Cells Smart Markers (C#)
// Description: Demonstrates how to bind a List<DataPoint> to a worksheet using Aspose.Cells smart markers, calculate the populated range, add a column chart, and set its series and category data programmatically before saving the file.
// Keywords: Aspose.Cells | C# | smart markers | dynamic chart | in‑memory list | bind collection to chart | NSeries range | Excel automation
// Common Searches: Aspose.Cells bind list to chart | C# create chart from collection using smart markers | set chart series range programmatically Aspose.Cells | dynamic Excel chart from POCO collection | update chart when data size changes Aspose.Cells
// Developer Intent: Populate an Excel chart directly from a C# collection without hard‑coding cell addresses.
// Use Cases: Generate a column chart that automatically reflects the items in a List<DataPoint>. | Adjust the chart range on‑the‑fly when the collection size varies. | Export a workbook with a ready‑to‑use chart for reporting dashboards.
// AI Prompts: Show how to switch the chart to a line type while keeping the smart‑marker data source. | Explain how to refresh the chart after adding new DataPoint objects at runtime. | Provide code to add a second series (e.g., a secondary value) from another property of the DataPoint class.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    // Simple POCO representing a data point
    // Demonstrates how to bind a List<DataPoint> to a worksheet using Aspose.Cells smart markers, calculate the populated range, add a column chart, and set its series and category data programmatically before saving the file.
    public class DataPoint
    {
        public string Category { get; set; }
        public double Value { get; set; }

        public DataPoint(string category, double value)
        {
            Category = category;
            Value = value;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare in‑memory data
            List<DataPoint> data = new List<DataPoint>
            {
                new DataPoint("A", 10),
                new DataPoint("B", 20),
                new DataPoint("C", 30),
                new DataPoint("D", 25)
            };

            // 2. Create a new workbook and add smart‑marker placeholders
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Smart‑marker rows – Aspose.Cells will expand these rows for each item in the list
            sheet.Cells["A2"].PutValue("&=Data.Category");
            sheet.Cells["B2"].PutValue("&=Data.Value");

            // 3. Bind the in‑memory list to the smart‑marker name "Data"
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process(); // Fills the worksheet with the list data

            // 4. Determine the last row that now contains data
            int lastRow = sheet.Cells.MaxDataRow; // zero‑based index
            // Convert to Excel row numbers (1‑based) for the range strings
            int firstDataRow = 2; // data starts at row 2 (A2:B2)
            string valueRange = $"=Sheet1!$B${firstDataRow}:$B${lastRow + 1}";
            string categoryRange = $"=Sheet1!$A${firstDataRow}:$A${lastRow + 1}";

            // 5. Add a chart and bind it to the filled ranges
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add(valueRange, true);          // Y‑values
            chart.NSeries.CategoryData = categoryRange;   // X‑axis categories
            chart.Title.Text = "Dynamic Chart from In‑Memory List";

            // 6. Save the workbook
            workbook.Save("DynamicChart.xlsx");
        }
    }
}
