using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace CorporateBranding
{
    public class TableStyleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the table
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Category");
                cells["C1"].PutValue("Price");
                for (int i = 2; i <= 6; i++)
                {
                    cells[$"A{i}"].PutValue($"Item{i - 1}");
                    cells[$"B{i}"].PutValue("General");
                    cells[$"C{i}"].PutValue(10 * i);
                }

                // -----------------------------------------------------------------
                // Create a custom table style that follows corporate branding guidelines
                // -----------------------------------------------------------------
                string corporateStyleName = "CorporateBrandStyle";

                // Add a new custom table style to the workbook's table style collection
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                int styleIndex = tableStyles.AddTableStyle(corporateStyleName);
                TableStyle corporateStyle = tableStyles[styleIndex];

                // Access the collection of style elements for the custom style
                TableStyleElementCollection elements = corporateStyle.TableStyleElements;

                // -------------------------------------------------
                // Header Row style (e.g., dark background, white bold text)
                // -------------------------------------------------
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.BackgroundColor = Color.FromArgb(0, 70, 127); // corporate dark blue
                headerStyle.Font.Color = Color.White;
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Size = 12;
                elements.Add(TableStyleElementType.HeaderRow);
                TableStyleElement headerElement = elements[TableStyleElementType.HeaderRow];
                headerElement.SetElementStyle(headerStyle);

                // -------------------------------------------------
                // First Column style (e.g., light gray background)
                // -------------------------------------------------
                Style firstColStyle = workbook.CreateStyle();
                firstColStyle.Pattern = BackgroundType.Solid;
                firstColStyle.BackgroundColor = Color.FromArgb(224, 224, 224); // light gray
                firstColStyle.Font.IsBold = true;
                elements.Add(TableStyleElementType.FirstColumn);
                TableStyleElement firstColElement = elements[TableStyleElementType.FirstColumn];
                firstColElement.SetElementStyle(firstColStyle);

                // -------------------------------------------------
                // Whole Table style (e.g., thin borders with corporate color)
                // -------------------------------------------------
                Style wholeTableStyle = workbook.CreateStyle();
                wholeTableStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                wholeTableStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                wholeTableStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                wholeTableStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                wholeTableStyle.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 70, 127);
                wholeTableStyle.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 70, 127);
                wholeTableStyle.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 70, 127);
                wholeTableStyle.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 70, 127);
                elements.Add(TableStyleElementType.WholeTable);
                TableStyleElement wholeTableElement = elements[TableStyleElementType.WholeTable];
                wholeTableElement.SetElementStyle(wholeTableStyle);

                // -------------------------------------------------
                // Create a ListObject (table) that uses the custom style
                // -------------------------------------------------
                int tableIndex = sheet.ListObjects.Add(0, 0, 5, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.TableStyleName = corporateStyleName;
                table.ShowTableStyleFirstColumn = true;   // ensure first column style is visible
                table.ShowTableStyleRowStripes = true;    // optional: add row stripes for readability

                // Save the workbook
                string outputPath = "CorporateBrandTableStyle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}