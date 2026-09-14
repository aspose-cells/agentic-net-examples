// Title: Group a Date column by month and year in an Aspose.Cells pivot table using C# to summarize sales data
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, inserts date and sales columns, adds a pivot table, and applies month‑year grouping to the Date row field. | Show how to apply month and year grouping to a pivot table field and then refresh and calculate the pivot using Aspose.Cells for .NET.
// Common Searches: aspnet aspose.cells group pivot table date by month and year c# example | c# create sales pivot table with month-year grouping using Aspose.Cells | how to group dates in an Aspose.Cells pivot table row field | refresh pivot table after grouping dates in Aspose.Cells .NET | Aspose.Cells pivot table grouping months years for sales report
// Tags: Aspose.Cells pivot table date grouping | C# month and year grouping in pivot table | sales data pivot table Aspose.Cells | Aspose.Cells refresh pivot calculation | create pivot table with Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds Date and Sales columns, builds a pivot table, groups the Date row field by months and years, refreshes and calculates the pivot, and saves the result as SalesByMonthYear.xlsx.
class GroupSalesByMonthYear
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add headers for the source data
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["B1"].PutValue("Sales");

        // Sample sales data with dates
        DateTime[] dates = new DateTime[]
        {
            new DateTime(2023, 1, 15),
            new DateTime(2023, 2, 20),
            new DateTime(2023, 3, 10),
            new DateTime(2023, 4, 5),
            new DateTime(2023, 5, 25),
            new DateTime(2023, 6, 12)
        };
        double[] sales = new double[] { 1500, 2300, 3200, 4100, 5000, 6000 };

        // Populate the worksheet with the sample data
        for (int i = 0; i < dates.Length; i++)
        {
            worksheet.Cells[i + 1, 0].PutValue(dates[i]); // Column A (Date)
            worksheet.Cells[i + 1, 1].PutValue(sales[i]); // Column B (Sales)
        }

        // Create a pivot table that covers the data range A1:B7 and place it at D3
        int pivotIndex = worksheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the Date field to the Row area and the Sales field to the Data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the PivotField representing the Date column
        PivotField dateField = pivotTable.RowFields[0];

        // Define grouping parameters: group by Months and Years for the whole year 2023
        DateTime startDate = new DateTime(2023, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);
        PivotGroupByType[] groupTypes = new PivotGroupByType[]
        {
            PivotGroupByType.Months,
            PivotGroupByType.Years
        };
        double interval = 1;          // Interval is required by the API (not used for date grouping)
        bool firstAsNewField = false; // Do not create a separate field for the first group

        // Apply the grouping to the date field
        dateField.GroupBy(startDate, endDate, groupTypes, interval, firstAsNewField);

        // Refresh and calculate the pivot table to reflect the grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("SalesByMonthYear.xlsx");
    }
}
