using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class HideEmptyRowsDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data (including some empty rows)
                dataSheet.Cells["A1"].Value = "Category";
                dataSheet.Cells["B1"].Value = "Product";
                dataSheet.Cells["C1"].Value = "Sales";

                dataSheet.Cells["A2"].Value = "Electronics";
                dataSheet.Cells["B2"].Value = "TV";
                dataSheet.Cells["C2"].Value = 1000;

                dataSheet.Cells["A3"].Value = "Electronics";
                dataSheet.Cells["B3"].Value = "";   // empty row
                dataSheet.Cells["C3"].Value = "";

                dataSheet.Cells["A4"].Value = "Furniture";
                dataSheet.Cells["B4"].Value = "Chair";
                dataSheet.Cells["C4"].Value = 500;

                dataSheet.Cells["A5"].Value = "Furniture";
                dataSheet.Cells["B5"].Value = "";   // empty row
                dataSheet.Cells["C5"].Value = "";

                // Add a pivot table based on the data range
                PivotTableCollection pivotTables = dataSheet.PivotTables;
                int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Hide empty rows in the pivot table.
                // The ShowEmptyRow property controls inclusion of empty rows.
                // Setting it to false hides empty rows.
                pivotTable.ShowEmptyRow = false;

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("HideEmptyRowsDemo.xlsx");
            }
            catch (Exception ex)
            {
                // Log or display the exception details
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}