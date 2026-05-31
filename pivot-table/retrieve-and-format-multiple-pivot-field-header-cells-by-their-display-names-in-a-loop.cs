using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RetrieveAndFormatPivotHeaders
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
                sheet.Cells["B1"].Value = "Product";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "North";
                sheet.Cells["B2"].Value = "Apple";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "North";
                sheet.Cells["B3"].Value = "Banana";
                sheet.Cells["C3"].Value = 800;

                sheet.Cells["A4"].Value = "South";
                sheet.Cells["B4"].Value = "Apple";
                sheet.Cells["C4"].Value = 1500;

                sheet.Cells["A5"].Value = "South";
                sheet.Cells["B5"].Value = "Banana";
                sheet.Cells["C5"].Value = 950;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Category, columns = Product, data = Sales (Sum)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

                // Refresh and calculate to generate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Format data field header (e.g., "Sum of Sales")
                for (int i = 0; i < pivotTable.DataFields.Count; i++)
                {
                    string displayName = pivotTable.DataFields[i].DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    if (headerCell != null)
                    {
                        Style style = headerCell.GetStyle();
                        style.Font.IsBold = true;
                        style.ForegroundColor = Color.Yellow;
                        style.Pattern = BackgroundType.Solid;
                        style.HorizontalAlignment = TextAlignmentType.Center;
                        headerCell.SetStyle(style);
                    }
                }

                // Format row field headers (Category)
                for (int i = 0; i < pivotTable.RowFields.Count; i++)
                {
                    string displayName = pivotTable.RowFields[i].DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    if (headerCell != null)
                    {
                        Style style = headerCell.GetStyle();
                        style.Font.IsBold = true;
                        style.ForegroundColor = Color.LightBlue;
                        style.Pattern = BackgroundType.Solid;
                        style.HorizontalAlignment = TextAlignmentType.Center;
                        headerCell.SetStyle(style);
                    }
                }

                // Save the workbook with the formatted pivot table
                string outputPath = "FormattedPivotHeaders.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveAndFormatPivotHeaders.Run();
        }
    }
}