// Title: Sort Excel rows by cell background color in column B using Aspose.Cells DataSorter (C#)
// Common Searches: C# Aspose.Cells sort worksheet rows by cell fill color | How to use DataSorter AddKey with SortOnType.CellColor in .NET | Sorting Excel data based on background color of a column using Aspose.Cells | Example of sorting Excel range by cell color with Aspose.Cells C# | Aspose.Cells sort rows by column B color ascending
// Tags: Aspose.Cells DataSorter cell color sort | C# Excel row ordering by background color | SortOnType.CellColor Aspose.Cells example | Excel column B color based sorting with Aspose | DataSorter AddKey SortOnType.CellColor usage

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, fills column B with distinct background colors, adds a DataSorter key on column B using SortOnType.CellColor, sorts the defined range in ascending order, and saves the result to 'SortedByCellColor.xlsx'.
class Program
{
    static void Main()
    {
        const string outputPath = "SortedByCellColor.xlsx";

        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data
            string[] items = { "Apple", "Banana", "Cherry", "Date", "Elderberry", "Fig", "Grape" };
            Color[] colors = {
                Color.LightGreen,
                Color.LightCoral,
                Color.LightBlue,
                Color.LightYellow,
                Color.LightGray,
                Color.LightPink,
                Color.LightSalmon
            };

            // Header
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Category");

            // Fill rows and apply background colors to column B
            for (int i = 0; i < items.Length; i++)
            {
                int row = i + 1; // data starts at row 2 (index 1)
                cells[row, 0].PutValue(items[i]);          // Column A
                cells[row, 1].PutValue($"Group {i + 1}"); // Column B (value not used for sorting)

                Style style = cells[row, 1].GetStyle();
                style.ForegroundColor = colors[i];          // Set background color directly
                style.Pattern = BackgroundType.Solid;
                cells[row, 1].SetStyle(style);
            }

            // Define the range that contains the data (including header)
            int startRow = 0;
            int startColumn = 0;
            int totalRows = items.Length + 1;
            int totalColumns = 2;

            // Configure the DataSorter to sort by the background color of column B (index 1)
            DataSorter sorter = workbook.DataSorter;
            // No ClearKeys method; simply add the required key
            sorter.AddKey(1, SortOnType.CellColor, SortOrder.Ascending, null);

            // Perform the sort
            sorter.Sort(cells, startRow, startColumn, totalRows, totalColumns);
            Console.WriteLine("Data sorted successfully based on cell background colors in column B.");

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
