using System.Drawing;
using Aspose.Cells;

class BackgroundColorSortExample
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a DataSorter instance
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // Assume the first row contains headers

        // Column U index (A=0, B=1, ..., U=20)
        int columnU = 20;

        // Define the color priority (highest priority first)
        Color[] priorityColors = new Color[]
        {
            Color.Red,
            Color.Orange,
            Color.Yellow,
            Color.Green,
            Color.Blue,
            Color.Purple
        };

        // Add a color key for each priority color
        foreach (Color clr in priorityColors)
        {
            sorter.AddColorKey(columnU, SortOnType.CellColor, SortOrder.Ascending, clr);
        }

        // Determine the range to sort (from the first row to the last used row in column U)
        int lastRow = worksheet.Cells.MaxDataRow;
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = columnU,
            EndRow = lastRow,
            EndColumn = columnU
        };

        // Perform the sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}