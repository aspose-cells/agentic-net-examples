using System;
using System.Drawing;
using Aspose.Cells;

class SortByBackgroundColor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a header
        worksheet.Cells["A1"].PutValue("Value");

        // Add sample data
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["A4"].PutValue(20);
        worksheet.Cells["A5"].PutValue(40);

        // Apply background colors to the data cells
        Style styleRed = workbook.CreateStyle();
        styleRed.ForegroundColor = Color.Red;
        styleRed.Pattern = BackgroundType.Solid;
        worksheet.Cells["A2"].SetStyle(styleRed);

        Style styleGreen = workbook.CreateStyle();
        styleGreen.ForegroundColor = Color.Green;
        styleGreen.Pattern = BackgroundType.Solid;
        worksheet.Cells["A3"].SetStyle(styleGreen);

        Style styleBlue = workbook.CreateStyle();
        styleBlue.ForegroundColor = Color.Blue;
        styleBlue.Pattern = BackgroundType.Solid;
        worksheet.Cells["A4"].SetStyle(styleBlue);

        Style styleYellow = workbook.CreateStyle();
        styleYellow.ForegroundColor = Color.Yellow;
        styleYellow.Pattern = BackgroundType.Solid;
        worksheet.Cells["A5"].SetStyle(styleYellow);

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row is a header

        // Add color sort keys (ascending order of colors)
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Red);
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Green);
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Blue);
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Yellow);

        // Define the range to sort (A1:A5)
        CellArea sortArea = CellArea.CreateCellArea("A1", "A5");

        // Perform the sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the workbook in XLSX format
        workbook.Save("SortedByBackgroundColor.xlsx");
    }
}