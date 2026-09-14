// Title: How to sort Excel rows by cell background color in column B using Aspose.Cells DataSorter (C#)
// AI Prompts: Generate C# code that creates a DataSorter, adds red, green, and blue color keys for column B, keeps the first row as a header, and sorts the range A1:B5 based on cell background colors. | Write a C# snippet that defines a custom color sequence, applies it to column B with Aspose.Cells, executes the sort, and prints each item with its associated color.
// Common Searches: C# Aspose.Cells sort rows by cell fill color | Aspose.Cells DataSorter AddColorKey example for custom color order | How to keep header row while sorting Excel sheet by background color using Aspose | Sorting Excel data by cell color in .NET with Aspose.Cells | Define multiple color keys for column sorting in Aspose.Cells C#
// Tags: Aspose.Cells DataSorter color key sorting | C# Excel color based row sorting | custom color order sorting Aspose.Cells | preserve header during Excel sort Aspose | Excel workbook cell fill color sort .NET

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsColorSortExample
{
    // The example creates a workbook, assigns red, green, and blue background colors to cells in column B, and uses Aspose.Cells' DataSorter with AddColorKey to define a custom color order. It preserves the header row, sorts the range A1:B5 by the cell colors, outputs the sorted items with their colors, and saves the result to an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Status");

            // Add sample data with different background colors in column B
            cells["A2"].PutValue("Task 1");
            cells["B2"].PutValue("Done");
            Style redStyle = workbook.CreateStyle();
            redStyle.ForegroundColor = Color.Red;
            redStyle.Pattern = BackgroundType.Solid;
            cells["B2"].SetStyle(redStyle);

            cells["A3"].PutValue("Task 2");
            cells["B3"].PutValue("In Progress");
            Style greenStyle = workbook.CreateStyle();
            greenStyle.ForegroundColor = Color.Green;
            greenStyle.Pattern = BackgroundType.Solid;
            cells["B3"].SetStyle(greenStyle);

            cells["A4"].PutValue("Task 3");
            cells["B4"].PutValue("Pending");
            Style blueStyle = workbook.CreateStyle();
            blueStyle.ForegroundColor = Color.Blue;
            blueStyle.Pattern = BackgroundType.Solid;
            cells["B4"].SetStyle(blueStyle);

            cells["A5"].PutValue("Task 4");
            cells["B5"].PutValue("Done");
            cells["B5"].SetStyle(redStyle);

            // Configure the DataSorter to sort by cell background color in column B (index 1)
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true; // First row contains headers

            // Define the order of colors: Red, Green, Blue (ascending)
            sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Red);
            sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Green);
            sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Blue);

            // Define the range to sort (A1:B5)
            CellArea sortArea = CellArea.CreateCellArea("A1", "B5");

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Output the sorted result to the console
            Console.WriteLine("Sorted data based on column B background colors:");
            for (int row = 1; row <= 4; row++) // rows 2-5 contain data
            {
                string item = cells[row, 0].StringValue;
                string status = cells[row, 1].StringValue;
                Color bgColor = cells[row, 1].GetStyle().ForegroundColor;
                Console.WriteLine($"{item} - {status} (Color: {bgColor.Name})");
            }

            // Save the workbook (optional)
            workbook.Save("ColorSortedOutput.xlsx");
        }
    }
}
