// Title: Aspose.Cells .NET – Show PivotChart Values as Percentage of Total in XLS
// Description: Load an existing XLS workbook, locate the first chart linked to a PivotTable, change the first data field to display percentages of the total, refresh the PivotTable and chart, and save the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PivotChart percentage | C# set pivot chart to percent of total | refresh pivot chart after data change | PivotField ShowValuesSetting PercentageOfTotal | convert pivot chart values to percent | Aspose.Cells XLS pivot table manipulation
// Common Searches: Aspose.Cells change pivot chart to percentage | C# code to set pivot chart values as percent of total | refresh pivot chart after modifying pivot table in .NET | show pivot table data as percentage using Aspose.Cells | how to display pivot chart percentages in XLS with C#
// Developer Intent: Modify a PivotChart so its data series are shown as percentages of the total and save the workbook.
// Use Cases: Create management reports where sales figures are displayed as a share of total revenue. | Update an existing KPI dashboard to automatically show ratio‑based metrics without rebuilding charts. | Batch‑process multiple Excel files to enforce consistent percentage formatting on all pivot charts.
// AI Prompts: Write C# code with Aspose.Cells that converts the first data field of a PivotChart to PercentageOfTotal and refreshes the chart. | Explain how to locate a PivotChart’s source PivotTable, set ShowValuesSetting.CalculationType, and save the workbook. | Provide error‑handling patterns for missing PivotChart or PivotTable when converting values to percentages in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartPercentage
{
    // Load an existing XLS workbook, locate the first chart linked to a PivotTable, change the first data field to display percentages of the total, refresh the PivotTable and chart, and save the updated file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "InputWorkbook.xls";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the first worksheet contains the pivot chart
                Worksheet sheet = workbook.Worksheets[0];

                // Find the first chart that is linked to a pivot table
                Chart pivotChart = null;
                foreach (Chart ch in sheet.Charts)
                {
                    if (!string.IsNullOrEmpty(ch.PivotSource))
                    {
                        pivotChart = ch;
                        break;
                    }
                }

                if (pivotChart == null)
                {
                    Console.WriteLine("No pivot chart found in the worksheet.");
                    return;
                }

                // Parse the PivotSource string: format "SheetName!PivotTableName"
                string[] parts = pivotChart.PivotSource.Split('!');
                if (parts.Length != 2)
                {
                    Console.WriteLine("Invalid PivotSource format.");
                    return;
                }

                string pivotSheetName = parts[0];
                string pivotTableName = parts[1];

                // Get the worksheet that holds the pivot table
                Worksheet pivotSheet = workbook.Worksheets[pivotSheetName];

                // Locate the pivot table by name
                PivotTable pivotTable = null;
                foreach (PivotTable pt in pivotSheet.PivotTables)
                {
                    if (pt.Name.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        pivotTable = pt;
                        break;
                    }
                }

                if (pivotTable == null)
                {
                    Console.WriteLine($"PivotTable '{pivotTableName}' not found.");
                    return;
                }

                // Set the data field to display values as percentage of total
                if (pivotTable.DataFields.Count > 0)
                {
                    PivotField dataField = pivotTable.DataFields[0];
                    dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
                }

                // Refresh pivot table data and recalculate
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Refresh the chart so it reflects the updated pivot data
                pivotChart.RefreshPivotData();

                // Output workbook path
                string outputPath = "OutputWorkbook.xls";

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
