using System;
using System.Drawing;
using Aspose.Cells;

class SortByCellColor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header row
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Status");

        // Row 2 - Red background
        cells["A2"].PutValue("Task1");
        cells["B2"].PutValue("Done");
        Style styleRed = workbook.CreateStyle();
        styleRed.ForegroundColor = Color.Red;
        styleRed.Pattern = BackgroundType.Solid;
        cells["B2"].SetStyle(styleRed);

        // Row 3 - Yellow background
        cells["A3"].PutValue("Task2");
        cells["B3"].PutValue("InProgress");
        Style styleYellow = workbook.CreateStyle();
        styleYellow.ForegroundColor = Color.Yellow;
        styleYellow.Pattern = BackgroundType.Solid;
        cells["B3"].SetStyle(styleYellow);

        // Row 4 - Green background
        cells["A4"].PutValue("Task3");
        cells["B4"].PutValue("Pending");
        Style styleGreen = workbook.CreateStyle();
        styleGreen.ForegroundColor = Color.Green;
        styleGreen.Pattern = BackgroundType.Solid;
        cells["B4"].SetStyle(styleGreen);

        // Configure the DataSorter to sort by cell background color in column B (index 1)
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row contains headers
        sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Red);
        // Additional color keys can be added if a specific order for other colors is required

        // Define the range to sort (including headers)
        CellArea sortArea = CellArea.CreateCellArea("A1", "B4");
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook
        workbook.Save("SortedByColor.xlsx");
    }
}