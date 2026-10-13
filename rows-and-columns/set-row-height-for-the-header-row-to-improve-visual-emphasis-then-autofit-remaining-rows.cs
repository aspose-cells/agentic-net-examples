// Title: Set header row height and auto‑fit remaining rows in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets the first worksheet row height to a specific point value, then iterates through the remaining rows to auto‑fit each one before saving the file. | Demonstrate how to emphasize a header row by adjusting its height and automatically resize all data rows in an Excel file using Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# set first row height to 30 points | How to auto‑fit rows after the header row using Aspose.Cells .NET | C# Aspose.Cells example for custom header height and row auto‑fit | Adjust Excel header row size and auto‑size data rows with Aspose.Cells library | Set row height and auto‑fit rows programmatically in Aspose.Cells for .NET
// Tags: set row height Aspose.Cells C# | auto‑fit rows Aspose.Cells .NET | header row formatting Excel Aspose | Aspose.Cells row height and autofit example | C# Excel workbook row sizing Aspose

using Aspose.Cells;
using System;

// Sets the first worksheet row height to 30 points, adds sample data, auto‑fits all subsequent rows, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the height of the header row (row index 0) to emphasize it
        sheet.Cells.SetRowHeight(0, 30); // Height in points

        // Example data population (optional, for demonstration)
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Row 1");
        sheet.Cells["A3"].PutValue("Row 2");
        sheet.Cells["B2"].PutValue(123);
        sheet.Cells["B3"].PutValue(456);

        // Auto‑fit all rows except the header row
        // MaxDataRow returns the last row that contains data
        for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
        {
            sheet.AutoFitRow(row);
        }

        // Save the workbook (lifecycle save)
        workbook.Save("output.xlsx");
    }
}
