// Title: C# – Create a Workbook, Add Sample Data, and Freeze the First Column with Aspose.Cells
// Description: Demonstrates how to instantiate a new Aspose.Cells Workbook, populate columns A and B with header and numeric rows, freeze column A using FreezePanes, and save the file as SampleFreezeFirstColumn.xlsx.
// Keywords: Aspose.Cells C# freeze column | FreezePanes Aspose.Cells example | create workbook Aspose.Cells | add sample data Excel C# | freeze first column Excel .NET | Aspose.Cells tutorial
// Common Searches: how to freeze the first column in Aspose.Cells C# | Aspose.Cells FreezePanes usage | C# code to create workbook and freeze header column | sample Aspose.Cells program with frozen column | freeze pane column A Aspose.Cells .NET
// Developer Intent: Generate a new Excel workbook, insert sample rows, lock the first column in place, and write the file to disk.
// Use Cases: Financial statements where item names in column A stay visible while scrolling horizontally. | Data‑entry templates that keep identifier columns fixed for easier navigation. | Large data exports where the primary key column must remain on screen.
// AI Prompts: Show how to freeze both rows and columns together with Aspose.Cells FreezePanes in C#. | Generate code that determines the number of header columns at runtime and freezes them automatically. | Explain each parameter of the FreezePanes method and how they control frozen rows and columns.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate a new Aspose.Cells Workbook, populate columns A and B with header and numeric rows, freeze column A using FreezePanes, and save the file as SampleFreezeFirstColumn.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data (header in column A and some data in columns A and B)
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Row 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Row 2");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("Row 3");
        worksheet.Cells["B4"].PutValue(300);

        // Freeze the first header column (column A)
        // Freeze at cell B1 with 0 frozen rows and 1 frozen column
        worksheet.FreezePanes("B1", 0, 1);

        // Save the workbook
        workbook.Save("SampleFreezeFirstColumn.xlsx");
    }
}
