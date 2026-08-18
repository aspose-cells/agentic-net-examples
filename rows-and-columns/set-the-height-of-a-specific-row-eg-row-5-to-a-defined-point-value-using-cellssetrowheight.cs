// Title: Aspose.Cells for .NET – Set a Specific Row Height (e.g., Row 5) in C#
// Description: Demonstrates how to create a workbook, access the first worksheet, and use worksheet.Cells.SetRowHeight(rowIndex, points) to set row 5 (zero‑based index 4) to 30 points, verify with GetRowHeight, and save as RowHeightDemo.xlsx.
// Keywords: Aspose.Cells set row height | C# Excel row height | SetRowHeight method | Aspose.Cells .NET row height example | adjust Excel row height programmatically | row 5 height Aspose | point value row height
// Common Searches: Aspose.Cells C# set row height to points | How to change height of row 5 in Excel using Aspose | SetRowHeight example .NET | Excel row height specific value Aspose.Cells | C# code to set Excel row height
// Developer Intent: Apply a precise point value to the height of a chosen row in an Excel worksheet using Aspose.Cells.
// Use Cases: Designing reports where header rows need extra vertical space for readability. | Standardizing row dimensions across generated worksheets before distribution. | Dynamically adjusting row heights based on content length or formatting rules.
// AI Prompts: Write C# code with Aspose.Cells that sets row 10 height to 25 points and saves the file as 'Report.xlsx'. | Explain how to read, modify, and batch‑apply row heights using Cells.SetRowHeight and Cells.GetRowHeight in Aspose.Cells. | Provide a loop example that sets rows 1‑20 to a uniform height of 20 points before exporting the workbook.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, access the first worksheet, and use worksheet.Cells.SetRowHeight(rowIndex, points) to set row 5 (zero‑based index 4) to 30 points, verify with GetRowHeight, and save as RowHeightDemo.xlsx.
class SetRowHeightExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the height of row 5 (zero‑based index 4) to 30 points
        worksheet.Cells.SetRowHeight(4, 30);

        // Verify the height (optional)
        Console.WriteLine("Row 5 height: " + worksheet.Cells.GetRowHeight(4));

        // Save the workbook
        workbook.Save("RowHeightDemo.xlsx");
    }
}
