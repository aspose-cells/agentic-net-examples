using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

class PivotChartPercentageExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the pivot table and the pivot chart
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (sheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables found in the first worksheet.");

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure the pivot table has at least one data field
            if (pivotTable.DataFields.Count == 0)
                throw new InvalidOperationException("Pivot table does not contain any data fields.");

            // Set the data field to display values as percentage of total
            PivotField dataField = pivotTable.DataFields[0];
            dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;

            // Refresh and recalculate the pivot table so the percentage values are applied
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Ensure a chart exists
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            // Get the first chart (assumed to be a pivot chart)
            Chart chart = sheet.Charts[0];

            // Link the chart to the pivot table
            chart.PivotSource = pivotTable.Name;

            // Optionally configure PivotOptions (e.g., enable drop zones)
            PivotOptions pivotOptions = chart.PivotOptions;
            pivotOptions.DropZonesVisible = true;

            // Refresh the chart data from the updated pivot table
            chart.RefreshPivotData();

            // Save the workbook with the updated pivot chart
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}