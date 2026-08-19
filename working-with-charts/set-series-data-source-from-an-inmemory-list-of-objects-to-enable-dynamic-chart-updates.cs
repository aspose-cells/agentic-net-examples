// Title: Create a Dynamic Column Chart in Aspose.Cells by Binding an In‑Memory List to Smart Markers
// Description: This example shows how to populate a worksheet from a C# List<DataPoint> using Aspose.Cells smart markers, process the markers with WorkbookDesigner, calculate the final row count, and bind the resulting cells to a column chart. The chart’s series and category ranges are set programmatically, and the workbook is saved as DynamicChart.xlsx.
// Keywords: Aspose.Cells | C# | .NET | smart markers | dynamic chart | bind list to worksheet | WorkbookDesigner | NSeries range | Excel chart from collection | in‑memory data source
// Common Searches: Aspose.Cells bind List<T> to chart | smart markers populate chart data .NET | set chart series range programmatically Aspose.Cells | dynamic Excel chart from in‑memory collection | calculate last row after smart marker processing
// Developer Intent: Bind an in‑memory collection to a worksheet with smart markers and use the expanded cells as the data source for a chart.
// Use Cases: Generate Excel reports where chart data reflects a runtime collection size. | Automatically adjust chart series when the underlying list grows or shrinks. | Create reusable templates that populate rows via smart markers and render charts without manual range updates.
// AI Prompts: Provide C# code that binds a List<DataPoint> to a worksheet using Aspose.Cells WorkbookDesigner and creates a column chart from the expanded cells. | Explain how to compute the last populated row after processing smart markers and set the NSeries values and category ranges dynamically. | Show how to change the chart type or add multiple series from additional in‑memory lists in the same workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    // Simple POCO representing a data point
    // This example shows how to populate a worksheet from a C# List<DataPoint> using Aspose.Cells smart markers, process the markers with WorkbookDesigner, calculate the final row count, and bind the resulting cells to a column chart. The chart’s series and category ranges are set programmatically, and the workbook is saved as DynamicChart.xlsx.
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

            // 2. Create a workbook and a worksheet that will hold the data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // 3. Insert smart markers – they will be replaced by the data source at processing time
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("&=$Data.Category"); // smart marker for Category
            dataSheet.Cells["B2"].PutValue("&=$Data.Value");    // smart marker for Value

            // 4. Bind the in‑memory list to the smart marker name "Data"
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process(); // expands the smart markers into actual rows

            // 5. Add a chart that uses the populated cells as its data source
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = dataSheet.Charts[chartIndex];

            // Determine the last row after processing (header + data count)
            int lastRow = data.Count + 1; // +1 because rows are 1‑based and header occupies row 1

            // Set the series range (values) and category axis range
            string valuesRange = $"=Sheet1!$B$2:$B${lastRow}";
            string categoryRange = $"=Sheet1!$A$2:$A${lastRow}";
            chart.NSeries.Add(valuesRange, true);
            chart.NSeries.CategoryData = categoryRange;

            // Optional: give the chart a title
            chart.Title.Text = "Dynamic Data Chart";

            // 6. Save the workbook
            workbook.Save("DynamicChart.xlsx");
        }
    }
}
