using System;
using System.Drawing;
using Aspose.Cells;

class SortByCellColor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Category");
        cells["C1"].PutValue("Value");

        // Add sample data with different background colors in column B
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue("Fruit");
        cells["C2"].PutValue(10);
        SetCellBackground(cells["B2"], Color.Red);

        cells["A3"].PutValue("Carrot");
        cells["B3"].PutValue("Vegetable");
        cells["C3"].PutValue(5);
        SetCellBackground(cells["B3"], Color.Green);

        cells["A4"].PutValue("Banana");
        cells["B4"].PutValue("Fruit");
        cells["C4"].PutValue(8);
        SetCellBackground(cells["B4"], Color.Blue);

        cells["A5"].PutValue("Broccoli");
        cells["B5"].PutValue("Vegetable");
        cells["C5"].PutValue(7);
        SetCellBackground(cells["B5"], Color.Red);

        // Configure the DataSorter to sort by cell background color in column B
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row is a header

        // Define custom color order: Red, then Green, then Blue
        sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Red);
        sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Green);
        sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Blue);

        // Define the range to sort (including header)
        CellArea area = CellArea.CreateCellArea("A1", "C5");

        // Perform the sort
        sorter.Sort(sheet.Cells, area);

        // Save the sorted workbook
        workbook.Save("SortedByCellColor.xlsx");
    }

    // Helper method to apply a solid background color to a cell
    static void SetCellBackground(Cell cell, Color color)
    {
        Style style = cell.GetStyle();
        style.ForegroundColor = color;
        style.Pattern = BackgroundType.Solid;
        cell.SetStyle(style);
    }
}