using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDateGrouping
{
    public class GroupByMonthsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a Sales column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 5));
            sheet.Cells["B2"].PutValue(1500);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 12));
            sheet.Cells["B3"].PutValue(2300);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 20));
            sheet.Cells["B4"].PutValue(3200);
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 8));
            sheet.Cells["B5"].PutValue(4100);
            sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 15));
            sheet.Cells["B6"].PutValue(5000);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area and Sales to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the PivotField that represents the Date column
            PivotField dateField = pivotTable.RowFields[0];

            // Define the grouping interval: months with an interval of 1 month
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate = new DateTime(2023, 12, 31);
            PivotGroupByType[] groupTypes = { PivotGroupByType.Months };
            double interval = 1;          // 1 month per group
            bool firstAsNewField = false; // Do not create a separate field for the first group

            // Apply grouping using the PivotField.GroupBy method (date overload)
            dateField.GroupBy(startDate, endDate, groupTypes, interval, firstAsNewField);

            // Refresh and calculate the pivot table to reflect the grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the grouped pivot table
            string outputPath = "PivotGroupedByMonths.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}