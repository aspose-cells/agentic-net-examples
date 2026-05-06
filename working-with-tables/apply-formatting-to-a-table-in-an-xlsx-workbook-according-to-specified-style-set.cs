using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using System.Drawing;

namespace AsposeCellsTableFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the table (A1:C5)
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Fruit");
            cells["C2"].PutValue(1.20);

            cells["A3"].PutValue("Carrot");
            cells["B3"].PutValue("Vegetable");
            cells["C3"].PutValue(0.80);

            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue("Fruit");
            cells["C4"].PutValue(1.10);

            cells["A5"].PutValue("Broccoli");
            cells["B5"].PutValue("Vegetable");
            cells["C5"].PutValue(1.50);

            // Add a ListObject (table) covering the data range
            int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Create a custom table style
            string customStyleName = "MyCustomStyle";
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
            int styleIdx = tableStyles.AddTableStyle(customStyleName);
            TableStyle customStyle = tableStyles[styleIdx];

            // Access the style elements collection
            TableStyleElementCollection elements = customStyle.TableStyleElements;

            // ----- Whole Table style -----
            int wholeIdx = elements.Add(TableStyleElementType.WholeTable);
            TableStyleElement wholeElement = elements[wholeIdx];
            Style wholeStyle = wholeElement.GetElementStyle();
            wholeStyle.Pattern = BackgroundType.Solid;
            wholeStyle.ForegroundColor = Color.LightYellow;
            wholeStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            wholeStyle.Borders[BorderType.BottomBorder].Color = Color.Gray;
            wholeElement.SetElementStyle(wholeStyle);

            // ----- Header Row style -----
            int headerIdx = elements.Add(TableStyleElementType.HeaderRow);
            TableStyleElement headerElement = elements[headerIdx];
            Style headerStyle = headerElement.GetElementStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerElement.SetElementStyle(headerStyle);

            // ----- First Column style -----
            int firstColIdx = elements.Add(TableStyleElementType.FirstColumn);
            TableStyleElement firstColElement = elements[firstColIdx];
            Style firstColStyle = firstColElement.GetElementStyle();
            firstColStyle.Font.IsBold = true;
            firstColStyle.Pattern = BackgroundType.Solid;
            firstColStyle.ForegroundColor = Color.LightCyan;
            firstColElement.SetElementStyle(firstColStyle);

            // Apply the custom style to the table
            table.ShowTableStyleFirstColumn = true;   // enable first column styling
            table.TableStyleName = customStyleName;

            // Additionally, apply the style to the entire table range using ApplyStyleToRange
            // (demonstrates the ListObject.ApplyStyleToRange method)
            table.ApplyStyleToRange();

            // Save the workbook
            workbook.Save("FormattedTable.xlsx", SaveFormat.Xlsx);
        }
    }
}