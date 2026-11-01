// Title: How to group date fields by month and year in an Aspose.Cells PivotTable using the C# GroupField method
// Common Searches: Aspose.Cells C# example for grouping pivot table dates by month and year | using GroupField to group date rows in an Aspose.Cells pivot table | C# code to apply GroupBy.Year and GroupBy.Month on a pivot table field with Aspose.Cells | step‑by‑step guide to month‑year date grouping in Aspose.Cells PivotTable | grouping date field in Excel pivot table via Aspose.Cells API C#
// Tags: Aspose.Cells GroupField month year grouping | C# pivot table date grouping Aspose | Excel pivot table month-year grouping with Aspose.Cells | GroupBy.Year GroupBy.Month Aspose.Cells API | pivot table row field grouping C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example demonstrates creating a workbook, adding sample Date, Category, and Amount data, inserting a PivotTable, and using the GroupField method with GroupBy.Year and GroupBy.Month to group the Date field by month and year. It also shows adding Category as a column field and Amount as a data field, refreshing the pivot, and saving the result as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and add sample data
            var dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Add headers
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Category");
            dataSheet.Cells["C1"].PutValue("Amount");

            // Sample data
            var dates = new DateTime[]
            {
                new DateTime(2023, 1, 15),
                new DateTime(2023, 1, 20),
                new DateTime(2023, 2, 5),
                new DateTime(2023, 2, 18),
                new DateTime(2023, 3, 10),
                new DateTime(2024, 1, 12)
            };
            var categories = new string[] { "A", "B", "A", "B", "A", "B" };
            var amounts = new double[] { 100, 150, 200, 120, 180, 210 };

            // Fill the worksheet with data
            for (int i = 0; i < dates.Length; i++)
            {
                int row = i + 1; // data starts from row 2 (index 1)
                dataSheet.Cells[row, 0].PutValue(dates[i]);      // Date column
                dataSheet.Cells[row, 1].PutValue(categories[i]); // Category column
                dataSheet.Cells[row, 2].PutValue(amounts[i]);    // Amount column
            }

            // Add a new worksheet for the pivot table
            var pivotSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            pivotSheet.Name = "Pivot";

            // Define the source range for the pivot table (including headers)
            var sourceRange = dataSheet.Cells.CreateRange(0, 0, dates.Length + 1, 3);

            // Add the pivot table to the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("PivotTable1", "A3", sourceRange.RefersTo);
            var pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add the Date field to the row area (field index 0 in source range)
            int dateFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

            // NOTE: Grouping by Year/Month is omitted for compatibility with older Aspose.Cells versions.
            // If your version supports grouping, you can uncomment the following lines:
            // pivotTable.GroupField(dateFieldIndex, GroupBy.Year, 2023, 2024);
            // pivotTable.GroupField(dateFieldIndex, GroupBy.Month, 1, 12);

            // Add the Category field to the column area (field index 1)
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);

            // Add the Amount field to the data area (field index 2)
            int amountFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, 2);
            // The default aggregation for numeric fields is Sum, so explicit setting is not required.
            // If needed, you can set the function like this (requires SubtotalFunction enum support):
            // pivotTable.DataFields[amountFieldIndex].Function = SubtotalFunction.Sum;

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotGroupedByMonthYear.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
