using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("InputFile.xls");

        // Access the worksheet that contains the pivot table (assumed first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Get the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Add a new chart (Column chart) to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Bind the chart to the pivot table using the PivotSource property
        // Format: SheetName!PivotTableName
        chart.PivotSource = $"{worksheet.Name}!{pivotTable.Name}";

        // Refresh the chart data from the pivot table
        chart.RefreshPivotData();

        // Refresh pivot tables in the worksheet to ensure data consistency
        worksheet.RefreshPivotTables();

        // Save the modified workbook
        workbook.Save("OutputFile.xlsx", SaveFormat.Xlsx);
    }
}