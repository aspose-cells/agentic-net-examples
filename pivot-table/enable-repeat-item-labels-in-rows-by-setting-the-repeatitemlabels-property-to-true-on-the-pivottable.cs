// Title: Enable Repeat Item Labels for Row Fields in Aspose.Cells PivotTable (C#)
// Description: Demonstrates how to set the IsRepeatItemLabels property to true for each row field in an Aspose.Cells PivotTable, refresh the pivot, calculate data, and save the workbook. Includes sample data creation, pivot configuration, and error handling.
// Keywords: Aspose.Cells repeat item labels | PivotTable IsRepeatItemLabels C# | repeat row labels Aspose.Cells | C# Aspose.Cells pivot settings | Enable repeat item labels | Aspose.Cells PivotTable example
// Common Searches: How to repeat row item labels in Aspose.Cells pivot table C# | Set IsRepeatItemLabels property for PivotField Aspose.Cells | Enable repeat item labels for pivot rows Aspose.Cells | Refresh pivot table after changing repeat item labels Aspose.Cells | Aspose.Cells C# repeat item labels tutorial
// Developer Intent: Set IsRepeatItemLabels = true on each row field of a PivotTable so that item labels are displayed on every row, then refresh and calculate the pivot.
// Use Cases: Create printable sales reports where each category label appears on every row for better readability. | Programmatically modify existing pivot tables to show repeated row labels before exporting to Excel. | Generate dynamic dashboards that require consistent row labeling across multiple data groups.
// AI Prompts: Write C# code using Aspose.Cells to enable repeat item labels for all row fields in a pivot table and save the workbook. | Explain the effect of the IsRepeatItemLabels property on pivot table layout in Aspose.Cells and how to refresh the table afterward. | Show how to toggle repeat item labels on specific row fields in an Aspose.Cells PivotTable for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to set the IsRepeatItemLabels property to true for each row field in an Aspose.Cells PivotTable, refresh the pivot, calculate data, and save the workbook. Includes sample data creation, pivot configuration, and error handling.
    public class EnableRepeatItemLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("SubCategory");
                worksheet.Cells["C1"].PutValue("Amount");

                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["B2"].PutValue("Apple");
                worksheet.Cells["C2"].PutValue(120);

                worksheet.Cells["A3"].PutValue("Fruit");
                worksheet.Cells["B3"].PutValue("Orange");
                worksheet.Cells["C3"].PutValue(150);

                worksheet.Cells["A4"].PutValue("Vegetable");
                worksheet.Cells["B4"].PutValue("Carrot");
                worksheet.Cells["C4"].PutValue(80);

                worksheet.Cells["A5"].PutValue("Vegetable");
                worksheet.Cells["B5"].PutValue("Broccoli");
                worksheet.Cells["C5"].PutValue(90);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table: Category as row, SubCategory as column, Amount as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Enable repeating item labels for each row field
                foreach (PivotField rowField in pivotTable.RowFields)
                {
                    rowField.IsRepeatItemLabels = true;
                }

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the modified pivot table
                workbook.Save("EnableRepeatItemLabelsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableRepeatItemLabelsDemo.Run();
        }
    }
}
