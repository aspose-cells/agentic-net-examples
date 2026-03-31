using System;
using System.Drawing;
using Aspose.Cells;

class SortByBackgroundColor
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "Input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // first row contains column headers

        // Sort rows based on the background color of column A (index 0)
        // Add custom color keys in the desired priority order: Red, Yellow, Green
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Red);
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Yellow);
        sorter.AddColorKey(0, SortOnType.CellColor, SortOrder.Ascending, Color.Green);

        // Define the range to sort (including headers)
        int lastRow = sheet.Cells.MaxDataRow;
        int lastCol = sheet.Cells.MaxDataColumn;
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = lastRow,
            EndColumn = lastCol
        };

        // Perform the sort
        sorter.Sort(sheet.Cells, sortArea);

        // Save the sorted workbook
        string outputPath = "SortedByColor.xlsx";
        workbook.Save(outputPath);
    }
}