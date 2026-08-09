// Title: Add a Line Chart to Every Worksheet and Link Data Labels to Cells – Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills each sheet with Category, Value, and Label columns, then adds a line chart to every worksheet. The chart’s series uses the Value column, the Category axis uses the Category column, and data labels are linked to the Label column so they update automatically. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells line chart each worksheet | C# link chart data labels to cells | batch create charts Aspose.Cells | ShowCellRange LinkedSource Aspose.Cells | populate worksheets sample data Aspose.Cells | Aspose.Cells ChartType.Line example | dynamic chart labels Excel API
// Common Searches: how to add a line chart to all sheets with Aspose.Cells | link chart data labels to a cell range in .NET | batch chart creation Aspose.Cells C# | set linked source for chart data labels Aspose.Cells | Aspose.Cells add line chart programmatically
// Developer Intent: Generate a line chart on each worksheet and bind its data‑label values to the corresponding cells.
// Use Cases: Automated sales dashboards where each month’s sheet shows a line chart with labels sourced from a description column. | Multi‑sheet financial reports that need identical line charts with dynamic labels that reflect cell edits. | Product performance workbooks where every sheet displays a line chart and the labels are driven by a separate label column for easy updates.
// AI Prompts: Write C# code using Aspose.Cells to add a bar chart to every worksheet and link its data labels to column D. | Show how to change the LinkedSource range of chart data labels after modifying the label column in an existing workbook. | Explain the purpose of ShowCellRange and LinkedSource properties when configuring chart series in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // C# example that creates a workbook, fills each sheet with Category, Value, and Label columns, then adds a line chart to every worksheet. The chart’s series uses the Value column, the Category axis uses the Category column, and data labels are linked to the Label column so they update automatically. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Populate each worksheet with sample data
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Header
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["C1"].PutValue("Label");

                    // Sample rows
                    for (int i = 2; i <= 6; i++)
                    {
                        sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                        sheet.Cells[$"B{i}"].PutValue(i * 10);
                        sheet.Cells[$"C{i}"].PutValue($"Lbl {i - 1}");
                    }
                }

                // Add a line chart to each worksheet and link data labels
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int lastRow = sheet.Cells.MaxDataRow; // zero‑based index
                    if (lastRow < 1) continue; // no data to chart

                    // Add a line chart
                    int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                    Chart chart = sheet.Charts[chartIndex];

                    // Build Excel‑style ranges
                    string sheetName = sheet.Name;
                    string valueRange = $"='{sheetName}'!$B$2:$B${lastRow + 1}";
                    string categoryRange = $"='{sheetName}'!$A$2:$A${lastRow + 1}";
                    string labelRange = $"='{sheetName}'!$C$2:$C${lastRow + 1}";

                    // Add series and set category data
                    chart.NSeries.Add(valueRange, true);
                    chart.NSeries.CategoryData = categoryRange;

                    // Configure data labels to show linked cell values
                    Series series = chart.NSeries[0];
                    series.DataLabels.ShowValue = true;
                    series.DataLabels.ShowCellRange = true;
                    series.DataLabels.LinkedSource = labelRange;
                }

                // Save the workbook
                string outputPath = "BatchLineCharts.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
