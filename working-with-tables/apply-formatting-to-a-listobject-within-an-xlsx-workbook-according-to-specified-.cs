using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using System.Drawing;

class ListObjectFormattingDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the table (A1:C4)
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Category");
        cells["C1"].PutValue("Price");

        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue("Fruit");
        cells["C2"].PutValue(1.20);

        cells["A3"].PutValue("Carrot");
        cells["B3"].PutValue("Vegetable");
        cells["C3"].PutValue(0.80);

        cells["A4"].PutValue("Bread");
        cells["B4"].PutValue("Bakery");
        cells["C4"].PutValue(2.50);

        // Add a ListObject (table) covering the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "C4", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Set a built‑in table style
        table.TableStyleType = TableStyleType.TableStyleMedium9;

        // Enable row and column stripe formatting
        table.ShowTableStyleRowStripes = true;
        table.ShowTableStyleColumnStripes = true;

        // Apply the table style to the underlying range
        table.ApplyStyleToRange();

        // Create a custom style for the "Price" column
        Style priceStyle = workbook.CreateStyle();
        priceStyle.Number = 10;                     // Currency format
        priceStyle.Font.Color = Color.DarkGreen;    // Font color
        priceStyle.ForegroundColor = Color.LightYellow;
        priceStyle.Pattern = BackgroundType.Solid;

        // Apply the custom style to the third column of the table
        table.ListColumns[2].SetDataStyle(priceStyle);

        // Save the workbook
        workbook.Save("FormattedListObject.xlsx", SaveFormat.Xlsx);
    }
}