// Title: How to configure Aspose.Cells C# pivot tables to show zeros for empty cells
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, enables the Worksheet.DisplayZeros option, adds data with missing entries, builds a pivot table, and saves the file so that missing values are rendered as 0. | Update an existing Aspose.Cells workbook to turn on zero display for empty cells and refresh its pivot tables, ensuring blanks appear as zero.
// Common Searches: Aspose.Cells C# show zeros for blank cells in a pivot table | Enable DisplayZeros on worksheet to affect pivot table values in Aspose .NET | How to make empty cells appear as 0 in Excel pivot tables using Aspose.Cells | C# Aspose.Cells pivot table zero handling for null data | Set Aspose.Cells workbook to display zero for missing values in pivot reports
// Tags: Worksheet.DisplayZeros option Aspose.Cells | pivot table zero handling Aspose.Cells | C# generate Excel pivot with Aspose.Cells | empty cell rendering as zero Aspose.Cells | configure Aspose.Cells workbook for zero display

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, sets Worksheet.DisplayZeros = true to ensure blanks are treated as zero, populates sample data with an empty cell, adds a pivot table on that range, configures row and data fields, calculates the pivot data, and saves the file, resulting in a pivot table where empty cells are displayed as 0.
    public class PivotTableDisplayZeroValuesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure zero values are displayed in the worksheet (affects pivot tables as well)
                sheet.DisplayZeros = true;

                // Populate sample data with some empty cells (null) that will be treated as zero
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("B");
                // B3 left empty – will be shown as zero
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(300);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D2", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableDisplayZeroValues.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
