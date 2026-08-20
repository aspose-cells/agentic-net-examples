// Title: C# – Create Workbook, Add Sample Data, and Freeze Header Row with Aspose.Cells
// Description: Demonstrates how to instantiate a Workbook, write a header and three data rows, freeze the first row using FreezePanes, and save the file as SampleDataFreezeHeader.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | FreezePanes example | create workbook Aspose.Cells | populate Excel cells C# | freeze first row Excel | save workbook Aspose.Cells | Excel automation .NET
// Common Searches: Aspose.Cells freeze first row C# | How to create and save a workbook with Aspose.Cells | Sample code for FreezePanes in Aspose.Cells | Write data to Excel and lock header using Aspose.Cells
// Developer Intent: Programmatically generate an Excel file, insert sample rows, lock the top header row, and write the workbook to disk.
// Use Cases: Building reports where column titles remain visible during scrolling. | Exporting database query results to Excel with a static header for end‑user analysis. | Creating reusable Excel templates that require the first row to stay fixed.
// AI Prompts: Generate a C# snippet that uses Aspose.Cells to create a workbook, add a header and data rows, freeze the first row, and save as .xlsx. | Explain the parameters of the FreezePanes method and show examples for freezing multiple rows and columns in Aspose.Cells. | Show how to apply style formatting to the header row before applying FreezePanes with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezeHeaderDemo
{
    // Demonstrates how to instantiate a Workbook, write a header and three data rows, freeze the first row using FreezePanes, and save the file as SampleDataFreezeHeader.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            // Header row
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Score");

            // Data rows
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["C2"].PutValue(85);

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");
            worksheet.Cells["C3"].PutValue(92);

            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue("Charlie");
            worksheet.Cells["C4"].PutValue(78);

            // Freeze the first header row (uses FreezePanes(string, int, int) rule)
            // Freeze at cell A2 with 1 frozen row and 0 frozen columns
            worksheet.FreezePanes("A2", 1, 0);

            // Save the workbook (uses Workbook.Save(string) rule)
            workbook.Save("SampleDataFreezeHeader.xlsx");

            Console.WriteLine("Workbook created, data populated, and header row frozen successfully.");
        }
    }
}
