// Title: Aspose.Cells for .NET – Set Row 10 Height to 30 Points (C# Example)
// Description: Demonstrates how to obtain the entire row range for row 10 in a worksheet, change its RowHeight property to 30 points, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set row height | entire row range Aspose.Cells | RowHeight property example | adjust row size .NET | Excel row height 30 points
// Common Searches: how to change row height with Aspose.Cells C# | retrieve entire row range Aspose.Cells | set specific row height in Excel using .NET | Aspose.Cells row height 30 points
// Developer Intent: Modify the height of row 10 by accessing its full row range and assigning a 30‑point value.
// Use Cases: Standardize header row dimensions for consistent report layouts. | Prepare worksheet rows for PDF conversion with fixed heights. | Ensure printable rows have uniform spacing across different page sizes.
// AI Prompts: Write C# code that sets rows 5‑15 to a height of 25 points with Aspose.Cells. | Show how to read the current RowHeight of a specific row before updating it. | Create an example that applies variable row heights based on the length of cell text using Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to obtain the entire row range for row 10 in a worksheet, change its RowHeight property to 30 points, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the entire row range for row 10 (rows are 1‑based in the address string)
            // Create a range for a single cell in row 10 and then use the EntireRow property
            AsposeRange entireRow = worksheet.Cells.CreateRange("A10").EntireRow;

            // Set the height of the retrieved row range to 30 points
            entireRow.RowHeight = 30;

            // Save the workbook
            workbook.Save("RowHeightDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
