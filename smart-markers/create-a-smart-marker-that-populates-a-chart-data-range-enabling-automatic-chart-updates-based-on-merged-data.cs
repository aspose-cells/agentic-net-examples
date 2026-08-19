// Title: C# – Populate Chart Data Range with Aspose.Cells Smart Markers and Auto‑Refresh the Chart
// Description: Demonstrates how to create a workbook, insert smart markers, bind a List<DataItem> to the marker name, process only the marker range, and automatically recalculate a column chart so it reflects the filled data. The example uses WorkbookDesigner, SetChartDataRange, and chart.Calculate in Aspose.Cells for .NET.
// Keywords: Aspose.Cells smart markers C# | auto update chart Aspose.Cells | WorkbookDesigner bind list | populate chart data range | chart.Calculate after smart marker processing | C# Excel chart automation
// Common Searches: Aspose.Cells smart marker chart example | C# fill Excel chart data with smart markers | How to refresh chart after WorkbookDesigner processing | SetChartDataRange smart markers Aspose.Cells | auto‑refresh Excel chart .NET
// Developer Intent: Create an Excel file where a smart‑marker range is filled from a .NET collection and the linked chart updates automatically without manual range changes.
// Use Cases: Generate a sales dashboard that inserts product categories and sales figures via smart markers and instantly updates a column chart. | Build a financial expense report where expense items are populated with smart markers and a pie chart reflects the distribution in real time. | Automate a quarterly performance workbook that binds a List<T> to smart markers and refreshes multiple charts after data insertion.
// AI Prompts: Show how to change the column chart to a line chart while keeping the smart‑marker data binding and automatic refresh. | Provide a sample that merges several smart‑marker ranges into one chart data range using Aspose.Cells. | Explain how to safely handle an empty data list so the chart does not reference invalid cells after processing.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerChartDemo
{
    // Simple data class for the smart marker data source
    // Demonstrates how to create a workbook, insert smart markers, bind a List<DataItem> to the marker name, process only the marker range, and automatically recalculate a column chart so it reflects the filled data. The example uses WorkbookDesigner, SetChartDataRange, and chart.Calculate in Aspose.Cells for .NET.
    public class DataItem
    {
        public string Category { get; set; } = string.Empty; // initialize to avoid nullable warning
        public double Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Set up headers ----------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // ---------- Insert smart markers for data rows ----------
                // These markers will be replaced by the data source values
                sheet.Cells["A2"].PutValue("&=$Data.Category");
                sheet.Cells["B2"].PutValue("&=$Data.Value");

                // Mark the range that contains smart markers (required for processing)
                AsposeRange smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // ---------- Add a chart that will use the populated data ----------
                // The chart is placed below the data area; its data range will be updated after processing
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Initially set a placeholder data range; it will be refreshed after smart markers are processed
                chart.SetChartDataRange("A2:B5", true); // A2:B5 will contain the filled data after processing

                // ---------- Prepare the data source ----------
                List<DataItem> data = new List<DataItem>
                {
                    new DataItem { Category = "Alpha",   Value = 120 },
                    new DataItem { Category = "Beta",    Value = 95  },
                    new DataItem { Category = "Gamma",   Value = 150 },
                    new DataItem { Category = "Delta",   Value = 80  }
                };

                // ---------- Configure WorkbookDesigner ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Bind the list to the smart marker name "Data"
                designer.SetDataSource("Data", data);

                // Process only the smart marker range (true = preserve unrecognized markers)
                designer.Process(true);

                // ---------- Recalculate the chart to reflect the new data ----------
                chart.Calculate();

                // ---------- Save the result ----------
                string outputPath = "SmartMarkerChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
