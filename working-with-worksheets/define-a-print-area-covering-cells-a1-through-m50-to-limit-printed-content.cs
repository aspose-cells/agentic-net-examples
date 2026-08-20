// Title: Set Print Area A1:M50 in an Aspose.Cells Workbook (C#)
// Description: This example creates a new workbook, fills cells A1‑M50 with sample data, assigns the range A1:M50 as the worksheet's print area using PageSetup.PrintArea, and saves the file. Only the defined range will be printed or exported.
// Keywords: Aspose.Cells C# print area | set worksheet print area Aspose.Cells | PageSetup.PrintArea example | define print range A1:M50 | limit printed content Aspose.Cells | Aspose.Cells for .NET print settings | C# workbook print area
// Common Searches: how to set print area in Aspose.Cells C# | Aspose.Cells define print area A1:M50 | programmatically set worksheet print range .NET | Aspose.Cells limit printed cells | C# set PageSetup.PrintArea
// Developer Intent: Configure a worksheet so that only cells A1 through M50 are included in print or export operations.
// Use Cases: Generating a report where only the first 50 rows and columns A‑M should appear on printed pages. | Creating invoice templates that automatically restrict the printable area to a predefined block. | Preparing a shared workbook that enforces a consistent print layout across all users.
// AI Prompts: Show C# code to set the print area to A1:M50 with Aspose.Cells and save the workbook. | Provide an Aspose.Cells example that populates a range, defines the print area, and exports the file to PDF. | Explain how to read the current PrintArea of a worksheet and modify it to a different range using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    // This example creates a new workbook, fills cells A1‑M50 with sample data, assigns the range A1:M50 as the worksheet's print area using PageSetup.PrintArea, and saves the file. Only the defined range will be printed or exported.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Populate some data within the print area for demonstration
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 13; col++) // Columns A (0) to M (12)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the print area covering cells A1 through M50
            worksheet.PageSetup.PrintArea = "A1:M50";

            // Save the workbook (the print area is stored in the file)
            workbook.Save("PrintAreaDemo.xlsx");

            Console.WriteLine("Workbook created with print area A1:M50.");
        }
    }
}
