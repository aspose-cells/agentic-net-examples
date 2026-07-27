using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add headers for first name, last name and a helper column (FullName)
        cells["A1"].PutValue("FirstName");
        cells["B1"].PutValue("LastName");
        cells["C1"].PutValue("FullName"); // helper column

        // Sample data: first name and last name pairs
        string[,] sampleData = {
            { "John", "Doe" },
            { "Alice", "Smith" },
            { "Bob", "Brown" },
            { "Charlie", "Adams" }
        };

        // Populate the worksheet with the sample data
        for (int i = 0; i < sampleData.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(sampleData[i, 0]); // FirstName
            cells[i + 1, 1].PutValue(sampleData[i, 1]); // LastName
        }

        // Compute the helper column by concatenating first and last names
        int dataRows = sampleData.GetLength(0);
        for (int i = 0; i < dataRows; i++)
        {
            string first = cells[i + 1, 0].StringValue;
            string last = cells[i + 1, 1].StringValue;
            cells[i + 1, 2].PutValue(first + " " + last); // FullName
        }

        // Configure the DataSorter to sort by the helper column (FullName)
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;          // First row contains headers
        sorter.Key1 = 2;                   // Column index of FullName (C)
        sorter.Order1 = SortOrder.Ascending;

        // Define the range to sort (including all three columns)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = dataRows,
            EndColumn = 2
        };

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Optionally hide the helper column after sorting
        worksheet.Cells.HideColumn(2);

        // Save the workbook
        workbook.Save("SortedByFullName.xlsx");
    }
}