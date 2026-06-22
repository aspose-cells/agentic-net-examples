using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableEnableMultipleSelectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Food";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["A3"].Value = "Drink";
                sheet.Cells["B3"].Value = 80;
                sheet.Cells["A4"].Value = "Snack";
                sheet.Cells["B4"].Value = 50;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table (Category as page field, Amount as data field)
                pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Enable multiple item selection for the page field
                if (pivotTable.PageFields.Count > 0)
                {
                    PivotField pageField = pivotTable.PageFields[0];
                    pageField.IsMultipleItemSelectionAllowed = true;
                }

                // Save the workbook to a file
                string outputPath = "PivotTable_EnableMultipleSelection.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableEnableMultipleSelectionDemo.Run();
        }
    }
}