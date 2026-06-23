using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class TableTotalRowAccent3Border
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Quantity");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(15);

                // Add a table with a total row
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowTotals = true;
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

                // Create a custom table style for the TotalRow element
                string customStyleName = "Accent3TotalRowStyle";
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                int styleIdx = tableStyles.AddTableStyle(customStyleName);
                TableStyle tableStyle = tableStyles[styleIdx];

                // Define style for TotalRow
                Style totalRowStyle = workbook.CreateStyle();
                totalRowStyle.Borders[BorderType.BottomBorder].ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                totalRowStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                totalRowStyle.Borders[BorderType.TopBorder].ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                totalRowStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                totalRowStyle.Borders[BorderType.LeftBorder].ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                totalRowStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                totalRowStyle.Borders[BorderType.RightBorder].ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                totalRowStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

                // Add TotalRow element to style and assign
                TableStyleElementCollection elements = tableStyle.TableStyleElements;
                int elementIdx = elements.Add(TableStyleElementType.TotalRow);
                elements[elementIdx].SetElementStyle(totalRowStyle);

                // Apply custom style to the table
                table.TableStyleName = customStyleName;

                // Save workbook
                string outputPath = "TableTotalRowAccent3Border.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            TableTotalRowAccent3Border.Run();
        }
    }
}