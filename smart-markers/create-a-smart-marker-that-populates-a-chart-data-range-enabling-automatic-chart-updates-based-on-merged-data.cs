// Title: C# – Populate an Excel Chart with Aspose.Cells Smart Markers and Auto‑Update the Series
// Description: Demonstrates how to create a workbook, place smart markers in a named range, bind a List<ChartDataItem> to the marker name, process only that range with WorkbookDesigner, and link a column chart to the generated data so the chart refreshes automatically.
// Keywords: Aspose.Cells | smart markers | chart data range | auto‑update chart | WorkbookDesigner | C# example | Excel chart generation | named range | list data source
// Common Searches: Aspose.Cells smart markers chart example C# | populate Excel chart using smart markers | auto refresh chart after WorkbookDesigner processing | bind list to smart markers Aspose.Cells | define smart marker range for chart data
// Developer Intent: Create an Excel file where a chart’s series is filled via smart markers, eliminating manual range adjustments.
// Use Cases: Generate sales‑by‑region charts from a collection of objects without hard‑coding the row count. | Build financial dashboards that expand or shrink automatically as new data rows are inserted. | Automate monthly reporting where chart visuals stay in sync with merged data rows produced by smart markers.
// AI Prompts: Show how to adjust the smart‑marker range for a dynamic number of rows at runtime. | Provide code to set the chart series to use only the value column while categories come from another column after processing. | Explain how to keep existing chart formatting intact when applying smart markers in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerChartDemo
{
    // Simple data class for the smart marker data source
    // Demonstrates how to create a workbook, place smart markers in a named range, bind a List<ChartDataItem> to the marker name, process only that range with WorkbookDesigner, and link a column chart to the generated data so the chart refreshes automatically.
    public class ChartDataItem
    {
        // Initialized with null-forgiving operator to satisfy non‑nullable warning
        public string Category { get; set; } = null!;
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

                // ---------- Prepare template with smart markers ----------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Smart marker rows (will be repeated for each data item)
                // The smart marker syntax "&=$Data.ColumnName" tells the designer to fill data from the source named "Data"
                sheet.Cells["A2"].PutValue("&=$Data.Category");
                sheet.Cells["B2"].PutValue("&=$Data.Value");

                // Define a named range that covers the smart marker rows.
                // This range will be processed by the designer.
                AsposeRange smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // ---------- Add a chart that references the data range ----------
                // The chart will be placed below the data (rows 5‑15, columns 0‑8)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Initially set the data range to the area that will be populated by smart markers.
                // After processing, the range will contain the actual data and the chart will update automatically.
                chart.NSeries.Add("=Sheet1!$A$2:$B$5", true);          // Values (both columns)
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";    // Categories (first column)

                // ---------- Prepare data source ----------
                List<ChartDataItem> data = new List<ChartDataItem>
                {
                    new ChartDataItem { Category = "Alpha",   Value = 120 },
                    new ChartDataItem { Category = "Beta",    Value = 95  },
                    new ChartDataItem { Category = "Gamma",   Value = 150 },
                    new ChartDataItem { Category = "Delta",   Value = 80  }
                };

                // ---------- Use WorkbookDesigner to process smart markers ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Bind the list to the smart marker name "Data"
                designer.SetDataSource("Data", data);
                // Process only the defined smart marker range (true = preserve unrecognized markers)
                designer.Process(smartMarkerRange, true);

                // ---------- Save the result ----------
                string outputPath = "SmartMarkerChartOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
