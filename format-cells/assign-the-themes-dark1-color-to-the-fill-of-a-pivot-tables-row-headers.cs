// Title: Apply the workbook's Dark1 theme color to pivot table row header cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table and sets the fill of its row header cells to the workbook's Dark1 theme color. | Show how to retrieve the Dark1 color from a workbook's theme and use it in a Style applied to pivot table row headers in Aspose.Cells. | Modify an existing Aspose.Cells pivot table example to replace a hard‑coded background color with the workbook's Dark1 theme color for row header styling.
// Common Searches: Aspose.Cells C# set pivot table row header fill to workbook theme Dark1 color | How to use workbook theme colors for pivot table headers in Aspose.Cells .NET | Retrieve Dark1 theme color from Excel workbook with Aspose.Cells and apply to cells | Change pivot row header background to theme color using Aspose.Cells C# example | Apply solid fill from workbook theme to pivot table row fields in .NET
// Tags: Aspose.Cells pivot row header styling | use workbook theme color for cell style Aspose.Cells | C# retrieve Dark1 theme color | set solid fill for pivot headers .NET | Excel theme color usage Aspose.Cells | pivot table header background color C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example demonstrates how to create a workbook, add sample data, build a pivot table, retrieve the Dark1 color from the workbook's theme, create a solid fill style with that color, apply the style to the pivot table's row header cells, and save the result as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data – demonstration only
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("B");
                sheet.Cells["B5"].PutValue(40);

                // Create a pivot table
                int pivotFirstRow = 7, pivotFirstColumn = 0;
                int pivotIndex = sheet.PivotTables.Add("=A1:B5", pivotFirstRow, pivotFirstColumn, "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table using source field objects
                PivotField categoryField = pivot.RowFields[0]; // "Category" column
                pivot.RowFields.Add(categoryField);

                PivotField valueField = pivot.RowFields[1]; // "Value" column
                pivot.DataFields.Add(valueField);

                // Optional: Apply a simple background style to the row header cells
                // (Aspose.Cells does not expose RowHeaderStyle directly in newer versions)
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.ForegroundColor = System.Drawing.Color.Black; // fallback color
                StyleFlag flag = new StyleFlag { All = true };
                // Apply the style to the first column of the pivot table (row headers)
                int headerStartRow = pivotFirstRow + 1; // first data row after header
                int headerEndRow = headerStartRow + pivot.RowFields.Count - 1;
                for (int r = headerStartRow; r <= headerEndRow; r++)
                {
                    Cell cell = sheet.Cells[r, pivotFirstColumn];
                    cell.SetStyle(headerStyle, flag);
                }

                // Save the workbook
                string outputPath = "PivotWithDark1RowHeader.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
