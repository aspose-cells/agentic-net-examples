// Title: Aspose.Cells for .NET: GetRowHeight and SetRowHeight to Transfer Row Height (C#)
// Description: Demonstrates how to create a workbook, assign a custom height to a source row, retrieve that height with GetRowHeight, and apply the same value to another row using SetRowHeight before saving the file.
// Keywords: Aspose.Cells | GetRowHeight | SetRowHeight | row height C# | copy row height Aspose | Excel row formatting .NET | retrieve row height programmatically | set row height example
// Common Searches: Aspose.Cells GetRowHeight example C# | How to set row height with Aspose.Cells | Copy row height from one row to another in Excel using .NET | Transfer row height between rows Aspose.Cells | SetRowHeight usage in C#
// Developer Intent: Read the height of a specific row and assign that exact measurement to another row in the same worksheet.
// Use Cases: Maintain uniform row spacing when duplicating template rows. | Adjust newly inserted rows to match the height of existing formatted rows. | Synchronize row heights across multiple sheets for consistent report layout.
// AI Prompts: Write C# code that uses Aspose.Cells to copy the height of row 5 to rows 10 through 15. | Explain the interaction between GetRowHeight and SetRowHeight for preserving row dimensions in an Excel file. | Provide a snippet that reads a row's height into a variable and applies it to several rows in a workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightDemo
{
    // Demonstrates how to create a workbook, assign a custom height to a source row, retrieve that height with GetRowHeight, and apply the same value to another row using SetRowHeight before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define source and destination row indices
            int sourceRowIndex = 0;      // Row 1 (zero‑based)
            int destinationRowIndex = 2; // Row 3 (zero‑based)

            // Set a custom height for the source row
            cells.SetRowHeight(sourceRowIndex, 30.0); // height in points

            // Retrieve the height of the source row using GetRowHeight
            double sourceHeight = cells.GetRowHeight(sourceRowIndex);
            Console.WriteLine($"Source row height (points): {sourceHeight}");

            // Explicitly set the same height to the destination row
            cells.SetRowHeight(destinationRowIndex, sourceHeight);
            Console.WriteLine($"Destination row height set to: {cells.GetRowHeight(destinationRowIndex)}");

            // Save the workbook (lifecycle save)
            workbook.Save("RowHeightCopyDemo.xlsx");
        }
    }
}
