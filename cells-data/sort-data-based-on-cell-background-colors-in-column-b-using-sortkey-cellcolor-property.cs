// Title: C# Aspose.Cells: Sort worksheet rows by column B cell background color using DataSorter
// Description: Demonstrates how to create a workbook, apply red, green, and blue background colors to column B cells, configure Aspose.Cells.DataSorter with AddColorKey (SortOnType.CellColor), define a range, sort rows according to the specified color order, output the result, and save the file as ColorSortedData.xlsx.
// Keywords: Aspose.Cells C# color sort | DataSorter AddColorKey example | Sort rows by cell background color .NET | Excel sort by cell color Aspose | SortOnType.CellColor usage | C# Aspose.Cells sorting tutorial | GitHub Aspose.Cells color sorting | global Excel automation | US .NET Excel library
// Common Searches: Aspose.Cells sort rows by cell color C# | How to use DataSorter AddColorKey in .NET | Excel background color sorting with Aspose | C# example sorting by column B cell color | SortOnType.CellColor Aspose.Cells tutorial
// Developer Intent: Arrange worksheet rows based on the background colors of cells in column B.
// Use Cases: Prioritize items in a product list where red‑marked rows must appear before green, blue, and uncolored rows. | Generate a status report that displays high‑priority (colored) entries first for easier review. | Prepare data for downstream processing where color‑coded categories dictate execution order.
// AI Prompts: Explain the role of DataSorter.AddColorKey and why the sequence of added keys determines the sorting hierarchy. | Show how to change the example to sort in descending order and add a custom Yellow color as the top priority. | Translate the C# color‑based sorting sample into equivalent VB.NET code using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsColorSortExample
{
    // Demonstrates how to create a workbook, apply red, green, and blue background colors to column B cells, configure Aspose.Cells.DataSorter with AddColorKey (SortOnType.CellColor), define a range, sort rows according to the specified color order, output the result, and save the file as ColorSortedData.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add header
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Category");

            // Add sample data
            cells["A2"].PutValue("Apple");
            cells["A3"].PutValue("Banana");
            cells["A4"].PutValue("Cherry");
            cells["A5"].PutValue("Date");

            // Apply background colors to column B cells
            Style styleRed = workbook.CreateStyle();
            styleRed.ForegroundColor = Color.Red;
            styleRed.Pattern = BackgroundType.Solid;
            cells["B2"].SetStyle(styleRed);

            Style styleGreen = workbook.CreateStyle();
            styleGreen.ForegroundColor = Color.Green;
            styleGreen.Pattern = BackgroundType.Solid;
            cells["B3"].SetStyle(styleGreen);

            Style styleBlue = workbook.CreateStyle();
            styleBlue.ForegroundColor = Color.Blue;
            styleBlue.Pattern = BackgroundType.Solid;
            cells["B4"].SetStyle(styleBlue);

            // Leave B5 without color (default)
            // Create a DataSorter and configure it to sort by cell background color in column B (index 1)
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true; // First row contains headers

            // Add color keys in the desired order (Red -> Green -> Blue -> No Color)
            sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Red);
            sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Green);
            sorter.AddColorKey(1, SortOnType.CellColor, SortOrder.Ascending, Color.Blue);

            // Define the range to sort (including header)
            CellArea sortArea = CellArea.CreateCellArea("A1", "B5");

            // Perform the sort
            sorter.Sort(worksheet.Cells, sortArea);

            // Output the sorted result to console for verification
            Console.WriteLine("Sorted data based on column B background colors:");
            for (int row = 1; row <= 5; row++)
            {
                string item = cells[row, 0].StringValue;
                string category = cells[row, 1].StringValue;
                Color bgColor = cells[row, 1].GetStyle().ForegroundColor;
                Console.WriteLine($"{item} - {category} (Color: {bgColor.Name})");
            }

            // Save the workbook
            workbook.Save("ColorSortedData.xlsx");
        }
    }
}
