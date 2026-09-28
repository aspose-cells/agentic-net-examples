// Title: Create a new Excel workbook, fill it with a 2‑D string array, and freeze the first header row using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to create a workbook, insert a two‑dimensional string array into the first worksheet, and apply FreezePanes to lock the top header row. | Show how to determine the populated range and call FreezePanes(1, 0, totalRows, totalCols) so the header stays visible while scrolling in an Aspose.Cells worksheet.
// Common Searches: Aspose.Cells C# freeze first row after writing data to worksheet | populate Excel worksheet from 2D string array using Aspose.Cells .NET | how to use FreezePanes with dynamic data range in Aspose.Cells | save workbook as xlsx after freezing header row with Aspose.Cells | example code for creating workbook and freezing header row in C# Aspose.Cells
// Tags: Aspose.Cells create workbook C# | Aspose.Cells write 2d array to worksheet | Aspose.Cells freeze header row | Aspose.Cells set FreezePanes range | Aspose.Cells save as xlsx

using Aspose.Cells;
using System;

// // Demonstrates creating a new workbook with Aspose.Cells, writing a 2‑D string array to the first sheet, freezing the top header row using FreezePanes, and saving the file as SampleData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data to populate
            string[,] data = new string[,]
            {
                { "ID", "Name", "Age", "Country" },
                { "1", "Alice", "30", "USA" },
                { "2", "Bob", "25", "UK" },
                { "3", "Charlie", "28", "Canada" },
                { "4", "Diana", "32", "Australia" }
            };

            // Fill the worksheet with the sample data
            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    sheet.Cells[row, col].PutValue(data[row, col]);
                }
            }

            // Freeze the first header row (row index 0)
            // Use the 4‑parameter overload: freeze rows above 'row' and columns left of 'column',
            // and set the scrollable area to include all populated cells.
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalCols = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(1, 0, totalRows, totalCols);

            // Save the workbook to a file
            workbook.Save("SampleData.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
