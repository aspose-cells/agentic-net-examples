using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    // Simple data model
    public class DataItem
    {
        public string Category { get; set; }
        public double Value { get; set; }

        public DataItem(string category, double value)
        {
            Category = category;
            Value = value;
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare in‑memory data
            List<DataItem> data = new List<DataItem>
            {
                new DataItem("A", 10),
                new DataItem("B", 20),
                new DataItem("C", 30),
                new DataItem("D", 25)
            };

            // Create a new workbook and add smart markers for the data source
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Smart markers – they will be replaced by the designer
            sheet.Cells["A2"].PutValue("&=$Data.Category");
            sheet.Cells["B2"].PutValue("&=$Data.Value");

            // Bind the in‑memory list to the smart marker name "Data"
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process(); // Populate the worksheet with the list values

            // Determine the last row that now contains data
            int lastRow = data.Count + 1; // +1 for header row

            // Add a chart that uses the populated cells as its data source
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the Y‑values range (Values) and the X‑axis (Category) range
            chart.NSeries.Add($"=Sheet1!$B$2:$B${lastRow}", true);
            chart.NSeries.CategoryData = $"=Sheet1!$A$2:$A${lastRow}";

            // Optional: give the chart a title
            chart.Title.Text = "Dynamic Chart from In‑Memory List";

            // Save the workbook
            workbook.Save("DynamicChart.xlsx");
        }
    }
}