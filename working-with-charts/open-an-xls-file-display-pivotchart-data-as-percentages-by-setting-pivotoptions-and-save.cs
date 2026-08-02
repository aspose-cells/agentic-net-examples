// Title: Show PivotChart Values as Percent of Total in XLS with Aspose.Cells (C#)
// Description: Loads an existing XLS workbook, finds the first chart linked to a pivot table, changes the first data field to display values as a percentage of the total, refreshes the pivot table and chart, and saves the result as a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# pivot chart percentage | PivotTable ShowValuesSetting PercentageOfTotal | Refresh pivot chart after data field change | Load and modify XLS workbook Aspose.Cells | C# Aspose.Cells PivotChart example
// Common Searches: Aspose.Cells set pivot chart to percent of total | C# change pivot table data field to percentage | Refresh pivot chart after updating calculation type Aspose | Find pivot chart source table in XLS using Aspose.Cells | Convert pivot chart values to percentages programmatically
// Developer Intent: Convert the data series of a PivotChart to show percentages of the total and save the updated XLS file.
// Use Cases: Create executive dashboards where sales figures are shown as share of total. | Automate updates to legacy XLS reports without rebuilding the workbook. | Generate periodic reports that require pivot chart percentages to reflect the latest data.
// AI Prompts: Generate C# code with Aspose.Cells that sets a pivot table’s first data field to PercentageOfTotal and updates the linked pivot chart. | Explain how to locate a pivot chart’s source pivot table in an XLS workbook using Aspose.Cells. | Provide error‑handling patterns for missing pivot charts or malformed PivotSource strings when converting chart data to percentages.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartPercentage
{
    // Loads an existing XLS workbook, finds the first chart linked to a pivot table, changes the first data field to display values as a percentage of the total, refreshes the pivot table and chart, and saves the result as a new file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load the existing XLS workbook
            Workbook workbook = new Workbook("input.xls");

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

            // Parse the PivotSource string (format: SheetName!PivotTableName)
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
                if (pt.Name == pivotTableName)
                {
                    pivotTable = pt;
                    break;
                }
            }

            if (pivotTable == null)
            {
                Console.WriteLine("Pivot table not found.");
                return;
            }

            // Set the data field to display values as percentage of total
            if (pivotTable.DataFields.Count > 0)
            {
                PivotField dataField = pivotTable.DataFields[0];
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
            }

            // Refresh and recalculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Refresh the chart so it reflects the updated pivot data
            pivotChart.RefreshPivotData();

            // Save the modified workbook
            workbook.Save("output.xls");
        }
    }
}
