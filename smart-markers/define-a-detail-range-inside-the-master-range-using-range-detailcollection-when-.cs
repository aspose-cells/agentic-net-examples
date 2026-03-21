using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Retrieve the master range by its predefined name (e.g., "MasterRange")
        // If the master range does not exist, the program exits gracefully.
        AsposeRange masterRange = workbook.Worksheets.GetRangeByName("MasterRange");
        if (masterRange == null)
        {
            Console.WriteLine("Master range 'MasterRange' not found.");
            return;
        }

        // Define the detail range inside the master range.
        // Here we offset the detail range by one row and one column from the master
        // and set its size to 3 rows x 3 columns (adjust as required).
        int detailFirstRow = masterRange.FirstRow + 1;
        int detailFirstColumn = masterRange.FirstColumn + 1;
        int detailRowCount = 3;
        int detailColumnCount = 3;

        // Create the detail range using the Cells.CreateRange method.
        AsposeRange detailRange = cells.CreateRange(detailFirstRow, detailFirstColumn, detailRowCount, detailColumnCount);

        // Optionally assign a name to the detail range for later reference.
        detailRange.Name = "DetailRange";

        // Add the newly created detail range to the worksheet's runtime range collection.
        cells.AddRange(detailRange);

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}