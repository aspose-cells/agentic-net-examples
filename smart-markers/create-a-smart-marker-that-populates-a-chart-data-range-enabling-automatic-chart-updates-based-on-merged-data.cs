// Title: Generate a column chart that auto‑updates via smart markers linked to a DataTable in Aspose.Cells for .NET
// AI Prompts: Write C# code that places SmartMarker:Category and SmartMarker:Value in worksheet cells, adds a column chart whose series formula references those smart markers, and processes the markers with WorkbookDesigner to produce an auto‑updating chart. | Show how to bind a DataTable to smart markers, invoke WorkbookDesigner.Process, and save the workbook as an Excel file containing a dynamically refreshed chart.
// Common Searches: Aspose.Cells C# smart marker syntax for chart series formula | How to bind a DataTable to smart markers and update an Excel chart automatically | Create a dynamic column chart from DataTable using WorkbookDesigner in Aspose.Cells | Auto‑refresh Excel chart when smart marker range changes in .NET | Smart markers for chart data range with merged cells Aspose.Cells example
// Tags: populate chart series with smart markers Aspose.Cells | WorkbookDesigner process smart markers C# | dynamic column chart from DataTable Aspose.Cells | auto‑updating Excel chart using smart markers | smart marker range for chart data Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, inserts SmartMarker:Category and SmartMarker:Value cells, adds a column chart whose series formula uses those smart markers, fills a DataTable with sample data, processes the smart markers with WorkbookDesigner, and saves the file as SmartMarkerChart.xlsx, resulting in a chart that updates automatically when the data changes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Set column headers
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Insert smart markers for the data rows (starting at row 2)
            sheet.Cells["A2"].PutValue("SmartMarker:Category");
            sheet.Cells["B2"].PutValue("SmartMarker:Value");

            // Add a column chart that will use the smart marker range as its data source
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIdx];
            // The series formula contains smart markers; after processing the chart will automatically reference the filled range
            chart.NSeries.Add("=Data!SmartMarker:Category,SmartMarker:Value", true);
            chart.Title.Text = "Sales by Category";

            // Prepare a data source (DataTable) that matches the smart marker fields
            DataTable dt = new DataTable();
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Value", typeof(double));
            dt.Rows.Add("North", 1200);
            dt.Rows.Add("South", 850);
            dt.Rows.Add("East", 970);
            dt.Rows.Add("West", 660);

            // Process the smart markers using WorkbookDesigner (correct API)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the workbook
            workbook.Save("SmartMarkerChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
