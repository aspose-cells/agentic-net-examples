using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    // Simple POCO representing a data item
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

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare in‑memory data source
            List<DataItem> items = new List<DataItem>
            {
                new DataItem("A", 10),
                new DataItem("B", 20),
                new DataItem("C", 30),
                new DataItem("D", 25)
            };

            // 2. Create a workbook template with smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Smart marker row – will be expanded by WorkbookDesigner
            sheet.Cells["A2"].PutValue("&=$Data.Category");
            sheet.Cells["B2"].PutValue("&=$Data.Value");

            // 3. Bind the in‑memory list to the smart marker name "Data"
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", items);
            designer.Process(); // Fills the smart marker rows with actual data

            // 4. Determine the data range after processing
            int firstDataRow = 2; // Excel rows are 1‑based; data starts at row 2
            int lastDataRow = firstDataRow + items.Count - 1;
            string valuesRange = $"=Sheet1!$B${firstDataRow}:$B${lastDataRow}";
            string categoryRange = $"=Sheet1!$A${firstDataRow}:$A${lastDataRow}";

            // 5. Add a chart and bind it to the populated range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series using the populated values range
            chart.NSeries.Add(valuesRange, true);
            // Set category (X‑axis) data
            chart.NSeries.CategoryData = categoryRange;

            // Optional: set a title
            chart.Title.Text = "Dynamic Data Chart";

            // 6. Save the workbook
            workbook.Save("DynamicChart.xlsx");
        }
    }
}