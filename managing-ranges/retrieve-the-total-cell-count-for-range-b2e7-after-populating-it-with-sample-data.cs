// Title: C# – Retrieve total cell count of range B2:E7 after populating data with Aspose.Cells
// Description: Creates a workbook, defines the B2:E7 range using Aspose.Cells.Range, fills each cell with sample values, calculates the total number of cells by multiplying RowCount and ColumnCount, prints the result, and saves the file as RangeCellCountDemo.xlsx.
// Keywords: Aspose.Cells C# range cell count | B2:E7 total cells | RowCount ColumnCount Aspose | populate range Aspose.Cells | calculate cells in range
// Common Searches: Aspose.Cells get number of cells in a range | C# count cells in B2:E7 using Aspose | RowCount * ColumnCount example Aspose.Cells | how to calculate total cells in a range .NET
// Developer Intent: Find out how to determine the total number of cells in the B2:E7 range after inserting sample data with Aspose.Cells for .NET.
// Use Cases: Validate that a generated data block occupies the expected cell count before exporting. | Drive progress bars or status updates while processing large worksheets. | Confirm that a dynamic range matches a predefined size for template integrity checks.
// AI Prompts: Write C# code that creates range B2:E7, fills it with sequential strings, and returns the total cell count using Aspose.Cells. | Explain the relationship between RowCount, ColumnCount, and total cells in an Aspose.Cells.Range. | Show an alternative way to obtain a range's cell count without manual multiplication in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines the B2:E7 range using Aspose.Cells.Range, fills each cell with sample values, calculates the total number of cells by multiplying RowCount and ColumnCount, prints the result, and saves the file as RangeCellCountDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create the range B2:E7 using Aspose.Cells.Range
            AsposeRange range = cells.CreateRange("B2", "E7");

            // Populate the range with sample data
            for (int i = 0; i < range.RowCount; i++)
            {
                for (int j = 0; j < range.ColumnCount; j++)
                {
                    // Example value: "R1C1", "R1C2", etc.
                    range[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Retrieve the total cell count for the range
            int totalCellCount = range.RowCount * range.ColumnCount;
            Console.WriteLine($"Total cells in range {range.Address}: {totalCellCount}");

            // Save the workbook
            workbook.Save("RangeCellCountDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
