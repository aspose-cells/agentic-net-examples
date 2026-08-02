// Title: Freeze the Header Row in a New Excel Workbook with Aspose.Cells for C#
// Description: Demonstrates how to create a workbook, add a header and sample rows, and lock the first row using Aspose.Cells' FreezePanes method before saving as SampleDataFreezeHeader.xlsx.
// Keywords: Aspose.Cells C# FreezePanes | freeze header row Excel | create workbook Aspose.Cells | sample data Excel C# | save Excel file Aspose.Cells
// Common Searches: Aspose.Cells freeze first row C# example | how to lock header row in Excel using Aspose.Cells | C# code to create workbook and freeze top row | FreezePanes usage Aspose.Cells .NET
// Developer Intent: Generate an Excel file, populate it with sample data, and keep the header row fixed while scrolling.
// Use Cases: Building reports where column titles remain visible on large sheets | Exporting data grids to Excel with a frozen top row for better navigation | Creating reusable templates that automatically apply a frozen header after data insertion
// AI Prompts: Write C# code that freezes multiple rows and columns based on user input using Aspose.Cells. | Show how to apply FreezePanes to a specific worksheet after dynamically adding data. | Explain how to unfreeze panes or modify the frozen range programmatically with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezeHeaderDemo
{
    // Demonstrates how to create a workbook, add a header and sample rows, and lock the first row using Aspose.Cells' FreezePanes method before saving as SampleDataFreezeHeader.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate header row
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Age");
            worksheet.Cells["D1"].PutValue("Score");

            // Populate some sample data rows
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["C2"].PutValue(28);
            worksheet.Cells["D2"].PutValue(85);

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Alice");
            worksheet.Cells["C3"].PutValue(32);
            worksheet.Cells["D3"].PutValue(92);

            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue("Bob");
            worksheet.Cells["C4"].PutValue(24);
            worksheet.Cells["D4"].PutValue(78);

            // Freeze the first header row (row 1) while keeping columns unfrozen
            // Freeze at cell A2, freeze 1 row, 0 columns
            worksheet.FreezePanes("A2", 1, 0);

            // Save the workbook
            workbook.Save("SampleDataFreezeHeader.xlsx");
        }
    }
}
