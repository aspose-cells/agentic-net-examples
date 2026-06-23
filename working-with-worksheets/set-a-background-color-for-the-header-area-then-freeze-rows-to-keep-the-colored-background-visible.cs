using System;
using Aspose.Cells;
using System.Drawing;

class SetHeaderBackgroundAndFreeze
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate header row and some sample data
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["C1"].PutValue("Header3");
        for (int i = 2; i <= 5; i++)
        {
            cells[$"A{i}"].PutValue($"Data {i - 1}");
            cells[$"B{i}"].PutValue(i * 10);
            cells[$"C{i}"].PutValue(i * 100);
        }

        // Create a style with a solid background color for the header
        Style headerStyle = workbook.CreateStyle();
        headerStyle.BackgroundColor = Color.LightBlue;   // Set desired background color
        headerStyle.Pattern = BackgroundType.Solid;      // Enable the background color

        // Apply the style to each cell in the first row (row index 0)
        for (int col = 0; col < 3; col++)
        {
            cells[0, col].SetStyle(headerStyle);
        }

        // Freeze the first row so the colored header remains visible while scrolling
        worksheet.FreezePanes("A2", 1, 0); // Freeze 1 row above cell A2

        // Save the workbook
        workbook.Save("HeaderBackgroundAndFreeze.xlsx");
    }
}