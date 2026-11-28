// Title: Create a line chart from an Excel template using Smart Markers and WorkbookDesigner in C# with Aspose.Cells
// Common Searches: how to generate a line chart from an Excel template with Aspose.Cells WorkbookDesigner C# | using smart markers to populate line chart data in Aspose.Cells C# example | add a line chart programmatically when the template worksheet has no chart Aspose.Cells | bind a DataTable to WorkbookDesigner smart markers Aspose.Cells C# | set chart data range for a line chart using Aspose.Cells C#
// Tags: WorkbookDesigner data source binding Aspose.Cells | Aspose.Cells generate line chart from template | set chart data range for line chart Aspose.Cells | process smart markers before chart creation Aspose.Cells | create line chart when worksheet has no charts Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SmartMarkerLineChartDemo
{
    // The example loads an Excel template containing Smart Markers, creates a DataTable with sales data, binds it to a WorkbookDesigner, processes the markers, and if the first worksheet lacks a chart, adds a line chart with revenue and cost series, defines the data range and titles, then saves the workbook as LineChartResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the template and the result workbook.
            // Ensure the template file exists in the same folder as the executable or provide a full path.
            string templatePath = "TemplateWithSmartMarkers.xlsx";
            string resultPath = "LineChartResult.xlsx";

            // Verify that the template file is present.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                return;
            }

            try
            {
                // ------------------------------------------------------------
                // 1. Prepare the data source (DataTable) that will be bound to
                //    the smart markers in the template.
                // ------------------------------------------------------------
                DataTable dt = new DataTable("SalesData");
                dt.Columns.Add("Month", typeof(string));
                dt.Columns.Add("Revenue", typeof(double));
                dt.Columns.Add("Cost", typeof(double));

                // Sample data – in a real scenario this could come from a DB, CSV, etc.
                dt.Rows.Add("Jan", 12000, 8000);
                dt.Rows.Add("Feb", 15000, 9000);
                dt.Rows.Add("Mar", 17000, 9500);
                dt.Rows.Add("Apr", 14000, 8500);
                dt.Rows.Add("May", 19000, 11000);
                dt.Rows.Add("Jun", 21000, 12000);

                // ------------------------------------------------------------
                // 2. Load the workbook template that contains smart markers.
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = new Workbook(templatePath);

                // ------------------------------------------------------------
                // 3. Bind the data source to the designer.
                //    The template uses the data source name "SalesData".
                // ------------------------------------------------------------
                designer.SetDataSource(dt); // By default the table name is used as the data source name.
                // If the template uses a different name you can bind like:
                // designer.SetDataSource("SalesData", dt);

                // ------------------------------------------------------------
                // 4. Process the smart markers.
                //    This will replace all markers (including those used for the chart)
                //    with the actual data values.
                // ------------------------------------------------------------
                designer.Process();

                // ------------------------------------------------------------
                // 5. (Optional) If the template does NOT already contain a chart,
                //    create a line chart programmatically using the processed data.
                // ------------------------------------------------------------
                Worksheet sheet = designer.Workbook.Worksheets[0];

                // Check whether a chart already exists; if not, add one.
                if (sheet.Charts.Count == 0)
                {
                    // Add a line chart positioned at rows 12‑30 and columns B‑H.
                    int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 1, 30, 7);
                    Chart lineChart = sheet.Charts[chartIndex];

                    // Define the data range for the chart.
                    // The range includes the header row (Month, Revenue, Cost) and the data rows.
                    // The "isVertical" flag set to true means series are plotted by column.
                    string dataRange = $"A1:C{dt.Rows.Count + 1}";
                    lineChart.SetChartDataRange(dataRange, true);

                    // Set titles (optional).
                    lineChart.Title.Text = "Revenue & Cost Over Months";
                    lineChart.NSeries.CategoryData = "A2:A7"; // Months as X‑axis categories.
                }

                // ------------------------------------------------------------
                // 6. Save the processed workbook.
                // ------------------------------------------------------------
                designer.Workbook.Save(resultPath, SaveFormat.Xlsx);
                Console.WriteLine($"Success: The line chart workbook has been saved to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred while generating the line chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
