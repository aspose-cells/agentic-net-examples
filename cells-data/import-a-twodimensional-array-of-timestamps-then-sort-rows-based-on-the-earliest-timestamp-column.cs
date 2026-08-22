// Title: Import a 2D DateTime array into an Aspose.Cells worksheet and sort rows by the earliest timestamp column (C#)
// AI Prompts: Write C# code that imports a two‑dimensional array of DateTime objects into an Aspose.Cells worksheet and then sorts the rows in ascending order based on the first column. | Show how to configure Aspose.Cells DataSorter to sort a range that contains timestamp values without a header row. | Modify the example to sort the data by the second timestamp column in descending order and save the workbook.
// Common Searches: c# aspocells import 2d datetime array and sort by earliest date | how to use DataSorter for timestamp columns in Aspose.Cells | sorting Excel rows by first date column using Aspose.Cells C# | importing multi‑column DateTime data into worksheet and ordering rows without headers
// Tags: import two‑dimensional DateTime array Aspose.Cells | DataSorter sort rows by first column | sort timestamps without header Aspose.Cells | C# sort worksheet range by earliest date | export sorted timestamps to Excel using Aspose.Cells

using System;
using Aspose.Cells;

// // Imports a 2D DateTime array into the first worksheet, defines the data range, uses DataSorter to order rows ascending by the earliest timestamp column, and saves the workbook as SortedTimestamps.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Two‑dimensional array of timestamps (rows = records, columns = different times)
        object[,] timestamps = new object[,]
        {
            { new DateTime(2023, 5, 10, 8, 30, 0), new DateTime(2023, 5, 10, 12, 0, 0) },
            { new DateTime(2023, 5, 9, 9, 15, 0),  new DateTime(2023, 5, 9, 13, 45, 0) },
            { new DateTime(2023, 5, 11, 7, 0, 0),  new DateTime(2023, 5, 11, 11, 30, 0) }
        };

        // Import the array into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportTwoDimensionArray(timestamps, 0, 0);

        // Determine the size of the imported range
        int rowCount = timestamps.GetLength(0);
        int colCount = timestamps.GetLength(1);

        // Configure the DataSorter to sort rows by the earliest timestamp (first column)
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = false;                     // No header row in this data set
        sorter.AddKey(0, SortOrder.Ascending);        // Sort by column A (index 0)

        // Define the area that contains the data to be sorted
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = rowCount - 1,
            EndColumn = colCount - 1
        };

        // Perform the sort operation
        sorter.Sort(cells, sortArea);

        // Save the workbook with the sorted data
        workbook.Save("SortedTimestamps.xlsx");
    }
}
