using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare data that will be placed in alternating rows
        // Each pair (string, int) represents a row's values
        object[] data = new object[]
        {
            "Item1", 100,
            "Item2", 200,
            "Item3", 300,
            "Item4", 400
        };

        // Import the data vertically starting at row 0, column 0
        // The 'skip' parameter is set to 1, which means one empty row will be left between each entry
        // Resulting rows: 0, 2, 4, 6 ... will contain the data; rows 1,3,5,... remain empty
        worksheet.Cells.ImportObjectArray(data, firstRow: 0, firstColumn: 0, isVertical: true, skip: 1);

        // Save the workbook to a file
        workbook.Save("AlternatingEntries.xlsx");
    }
}