using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsExamples
{
    public class DeleteRowFromListObjectAndVerifyChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (including header) for the table
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Series1");
                cells["C1"].PutValue("Series2");

                cells["A2"].PutValue("Q1");
                cells["B2"].PutValue(100);
                cells["C2"].PutValue(150);

                cells["A3"].PutValue("Q2");
                cells["B3"].PutValue(200);
                cells["C3"].PutValue(250);

                cells["A4"].PutValue("Q3");
                cells["B4"].PutValue(300);
                cells["C4"].PutValue(350);

                // Create a ListObject (table) that covers the data range A1:C4
                int tableIndex = sheet.ListObjects.Add(0, 0, 3, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesData";

                // Add a column chart that uses the table data
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the series using the table's column ranges
                chart.NSeries.Add(table.DisplayName + "[Series1]", true);
                chart.NSeries.Add(table.DisplayName + "[Series2]", true);
                chart.NSeries.CategoryData = table.DisplayName + "[Category]";

                // Output initial state
                Console.WriteLine("Initial chart series count: " + chart.NSeries.Count);
                Console.WriteLine("Initial table row count: " + (table.DataRange.RowCount - 1));

                // Delete the second data row (row index 2 in worksheet, zero‑based)
                cells.DeleteRow(2, true); // true updates references

                // After deletion, the table automatically adjusts its range and rows
                Console.WriteLine("\nAfter deleting a row from the worksheet:");
                Console.WriteLine("Updated chart series count: " + chart.NSeries.Count);
                Console.WriteLine("Updated table row count: " + (table.DataRange.RowCount - 1));

                // Verify that the chart still has the same number of series
                if (chart.NSeries.Count == 2)
                {
                    Console.WriteLine("Chart series count is correct after row deletion.");
                }
                else
                {
                    Console.WriteLine("Unexpected chart series count after row deletion.");
                }

                // Save the workbook
                string outputPath = "DeleteRowFromListObjectAndVerifyChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("\nWorkbook saved to: " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point required for console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            DeleteRowFromListObjectAndVerifyChart.Run();
        }
    }
}