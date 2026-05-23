using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

namespace AsposeCellsPivotThemeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
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

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Configure the pivot table: add row field and data field
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Ensure the style is applied to row headers
                pivot.ShowPivotStyleRowHeader = true;

                // Calculate the pivot data so that the layout is generated
                pivot.CalculateData();

                // Create a style that uses the theme's Background1 (Dark1) color as fill
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Background1, 0);
                headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0);
                headerStyle.Font.IsBold = true;

                // Apply the style to the row header cells.
                // Use the pivot table's overall range (TableRange1) to locate the first column.
                CellArea tableArea = pivot.TableRange1;
                int headerColumn = tableArea.StartColumn; // first column of the pivot table
                // Row headers start after the column header row (offset by 1)
                for (int r = tableArea.StartRow + 1; r <= tableArea.EndRow; r++)
                {
                    pivot.Format(r, headerColumn, headerStyle);
                }

                // Save the workbook
                string outputPath = "PivotTableRowHeaderThemeDark1.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}