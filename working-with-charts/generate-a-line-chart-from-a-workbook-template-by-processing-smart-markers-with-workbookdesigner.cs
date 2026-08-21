// Title: Create a Line Chart from a Smart‑Marker Template with WorkbookDesigner in Aspose.Cells for .NET
// Description: Load a workbook template that contains smart markers, bind a JSON data source, process the markers with WorkbookDesigner, add a line chart that references the populated cells, set title and legend, and save the workbook as a new Excel file.
// Keywords: Aspose.Cells line chart | WorkbookDesigner smart markers | C# JSON data source Aspose.Cells | add chart programmatically .NET | set chart data range Aspose.Cells | smart marker template Excel | automate chart generation C#
// Common Searches: Aspose.Cells create line chart after processing smart markers | WorkbookDesigner JSON data source example | C# add line chart to Excel with Aspose.Cells | set dynamic chart range Aspose.Cells | smart markers line chart template
// Developer Intent: Generate a line chart in an Excel workbook by processing smart markers with WorkbookDesigner and binding JSON data.
// Use Cases: Automatically populate a sales‑report template with JSON data and produce a monthly sales line chart. | Reuse a single chart template to create multiple workbooks, each showing trend lines for different data sets. | Integrate API‑driven JSON responses into Excel and generate performance trend charts without manual editing.
// AI Prompts: Show how to calculate the chart data range dynamically based on the number of rows in the JSON array. | Provide code to add data labels and customize line colors and markers for the generated chart. | Explain how to reference a named range instead of a hard‑coded A2:B6 range for the chart source.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSmartMarkerLineChart
{
    // Load a workbook template that contains smart markers, bind a JSON data source, process the markers with WorkbookDesigner, add a line chart that references the populated cells, set title and legend, and save the workbook as a new Excel file.
    class Program
    {
        static void Main()
        {
            // Load the workbook template that contains smart markers.
            // The template should have a range named "_CellsSmartMarkers" or use line‑by‑line processing.
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Example JSON data source that matches the smart markers in the template.
            // Adjust the JSON structure to correspond to the markers you placed in the template.
            string jsonData = @"{
                ""SalesData"": [
                    { ""Month"": ""Jan"", ""Amount"": 1200 },
                    { ""Month"": ""Feb"", ""Amount"": 1500 },
                    { ""Month"": ""Mar"", ""Amount"": 1800 },
                    { ""Month"": ""Apr"", ""Amount"": 2100 },
                    { ""Month"": ""May"", ""Amount"": 2400 }
                ]
            }";

            // Bind the JSON data to a smart‑marker name (e.g., "SalesData").
            designer.SetJsonDataSource("SalesData", jsonData);

            // Process the smart markers and populate the worksheet with the data.
            designer.Process();

            // After processing, add a line chart that visualizes the populated data.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line chart to the worksheet. Position it from row 10, column 1 to row 30, column 10.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 9, 0, 29, 9);
            Chart lineChart = sheet.Charts[chartIndex];

            // Define the data range for the chart.
            // Assuming the processed data starts at A2 (Month) and B2 (Amount) and extends downwards.
            // Adjust the range as needed based on the actual data size.
            lineChart.SetChartDataRange("A2:B6", true);

            // Optional: set chart title and enable legend.
            lineChart.Title.Text = "Monthly Sales";
            lineChart.ShowLegend = true;

            // Save the resulting workbook with the chart.
            workbook.Save("OutputWithLineChart.xlsx");
        }
    }
}
