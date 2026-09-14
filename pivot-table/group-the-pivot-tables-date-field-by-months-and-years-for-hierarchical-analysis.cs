// Title: How to group a PivotTable date field by years and months using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that creates a pivot table from worksheet data and groups the Date row field into years and months with Aspose.Cells. | Write code to refresh the pivot cache and recalculate the pivot table after applying hierarchical date grouping in Aspose.Cells. | Show how to save the workbook as an .xlsx file after grouping the Date field by year and month using Aspose.Cells for .NET.
// Common Searches: aspnet c# group pivot table date column by month and year using Aspose.Cells library | example code for hierarchical date grouping in Excel pivot tables with Aspose.Cells for .NET | how to refresh and calculate pivot table after grouping dates in Aspose.Cells C# | Aspose.Cells pivot table grouping years then months programmatically | save pivot table with grouped date fields to Excel file using Aspose.Cells C#
// Tags: Aspose.Cells pivot table date grouping | group dates by years and months Aspose.Cells | refresh pivot cache Aspose.Cells C# | calculate pivot data Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotDateGroupingDemo
{
    // Demonstrates creating a workbook with sample dates and sales, building a pivot table, grouping the Date row field by years and months, refreshing and recalculating the pivot, and saving the result as an Excel .xlsx file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                // Header row
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Sales");

                // Sample dates spanning several months and years
                DateTime[] dates = new DateTime[]
                {
                    new DateTime(2022, 1, 15),
                    new DateTime(2022, 2, 10),
                    new DateTime(2022, 3, 5),
                    new DateTime(2023, 1, 20),
                    new DateTime(2023, 2, 25),
                    new DateTime(2023, 3, 30)
                };

                // Corresponding sales values
                double[] sales = new double[] { 1200, 1500, 1800, 2000, 2300, 2600 };

                // Fill the worksheet with the data
                for (int i = 0; i < dates.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(dates[i]); // Column A
                    sheet.Cells[i + 2, 1].PutValue(sales[i]); // Column B
                }

                // ---------- Create a pivot table ----------
                // Data range includes the header row and all data rows
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the Date field to the Row area
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // 0 = first column (Date)

                // Add the Sales field to the Data area
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // 1 = second column (Sales)

                // ---------- Group the Date field by Years and Months ----------
                // Retrieve the pivot field that represents the Date column
                PivotField dateField = pivotTable.RowFields[0];

                // Define the grouping range (cover all dates in the source data)
                DateTime startDate = new DateTime(2022, 1, 1);
                DateTime endDate = new DateTime(2023, 12, 31);

                // Specify the grouping types: Years first, then Months
                PivotGroupByType[] groupTypes = new PivotGroupByType[]
                {
                    PivotGroupByType.Years,
                    PivotGroupByType.Months
                };

                // Perform the grouping; 'false' means the first group (Years) will replace the original field
                dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

                // Optional: Verify that grouping was applied
                if (dateField.GroupSettings != null)
                {
                    Console.WriteLine("Date field has been grouped.");
                }

                // Refresh the pivot cache and calculate the pivot table to apply the grouping
                pivotTable.RefreshData();          // Correct method to refresh cache
                pivotTable.CalculateData();        // Recalculate pivot data

                // ---------- Save the workbook ----------
                workbook.Save("PivotDateGroupedByMonthsAndYears.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
