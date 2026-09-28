// Title: How to group a PivotTable date field by month interval in Aspose.Cells for .NET (C#)
// AI Prompts: Create a PivotTable, add a date field to the row area, and group it by months with a 1‑month interval using Aspose.Cells in C#. | Apply month‑based grouping to an existing date pivot field without creating a new field, then refresh the pivot cache. | Calculate the PivotTable after month grouping of the OrderDate field and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# group pivot table date field by month interval | example of month grouping in Aspose.Cells PivotTable using C# | how to set month grouping on a date field in an Aspose.Cells pivot table | refresh pivot cache after grouping dates with Aspose.Cells .NET
// Tags: Aspose.Cells PivotField.GroupBy method | C# month-based grouping for pivot tables | refresh pivot cache Aspose.Cells | save workbook as xlsx using Aspose.Cells | date row field grouping Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook with sample order dates and sales, adds a PivotTable, places the OrderDate field in the row area, groups this date field by months using a 1‑month interval, refreshes and calculates the PivotTable, and saves the result as PivotDateFieldMonthGrouping.xlsx.
    public class PivotDateFieldMonthGroupingDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Populate sample data -----
                // Header row
                sheet.Cells["A1"].Value = "OrderDate";
                sheet.Cells["B1"].Value = "Sales";

                // Sample dates (spread across several months) and sales values
                sheet.Cells["A2"].Value = new DateTime(2023, 1, 15);
                sheet.Cells["B2"].Value = 1500;

                sheet.Cells["A3"].Value = new DateTime(2023, 2, 20);
                sheet.Cells["B3"].Value = 2300;

                sheet.Cells["A4"].Value = new DateTime(2023, 3, 10);
                sheet.Cells["B4"].Value = 3200;

                sheet.Cells["A5"].Value = new DateTime(2023, 4, 5);
                sheet.Cells["B5"].Value = 4100;

                sheet.Cells["A6"].Value = new DateTime(2023, 5, 25);
                sheet.Cells["B6"].Value = 5000;

                // ----- Create a pivot table -----
                // Data range A1:B6, place pivot table at C3, name it "SalesPivot"
                int pivotIndex = sheet.PivotTables.Add("A1:B6", "C3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the date field to the row area and the sales field to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "OrderDate");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // ----- Group the date field by months -----
                // Retrieve the date pivot field (first row field)
                PivotField dateField = pivotTable.RowFields[0];

                // Define the grouping range (cover all dates in the source data)
                DateTime startDate = new DateTime(2023, 1, 1);
                DateTime endDate = new DateTime(2023, 12, 31);

                // Specify that we want to group by months
                PivotGroupByType[] groupTypes = new PivotGroupByType[] { PivotGroupByType.Months };

                // Interval of 1 month, do not create a new field (group in place)
                double interval = 1.0;
                bool firstAsNewField = false;

                // Apply grouping using the PivotField.GroupBy overload for date fields
                dateField.GroupBy(startDate, endDate, groupTypes, interval, firstAsNewField);

                // Refresh the pivot cache and calculate the pivot table to reflect the grouping
                pivotTable.RefreshData();      // Correct method to refresh cache
                pivotTable.CalculateData();

                // ----- Save the workbook -----
                workbook.Save("PivotDateFieldMonthGrouping.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
