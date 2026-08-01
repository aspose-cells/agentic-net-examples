// Title: C# – Add a Running Total Column with SUM Formula Using Aspose.Cells for .NET
// Description: This Aspose.Cells example creates a workbook, inserts a new column, and fills it with cumulative totals by applying a SUM formula that references the first amount cell ($B$2) through each row, then saves the file as RunningTotalExample.xlsx.
// Keywords: Aspose.Cells | running total column | cumulative sum | SUM formula | C# Excel automation | insert column programmatically | Excel macro‑style routine | calculate running total | Aspose.Cells .NET example | Excel workbook generation
// Common Searches: how to add a cumulative total column with Aspose.Cells C# | Aspose.Cells insert column and set SUM formula across rows | C# code for running total in Excel using Aspose.Cells | programmatically create running balance column in .xlsx | Aspose.Cells example for cumulative sales totals
// Developer Intent: Insert a new column and populate each cell with a running total calculated via a SUM formula.
// Use Cases: Generate a sales ledger where each row shows the cumulative revenue. | Create a budgeting sheet that tracks progressive expense totals. | Automate invoice workbooks to display a running balance after each line item.
// AI Prompts: Write Aspose.Cells C# code that adds a running total column using an absolute start reference and a relative end reference for each row. | Modify the routine to automatically detect the last data row instead of using a hard‑coded index. | Provide an alternative implementation that uses the SUBTOTAL function for the running total in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRunningTotalExample
{
    // This Aspose.Cells example creates a workbook, inserts a new column, and fills it with cumulative totals by applying a SUM formula that references the first amount cell ($B$2) through each row, then saves the file as RunningTotalExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (e.g., sales amounts in column B)
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("Item 1");
                cells["B2"].PutValue(100);
                cells["A3"].PutValue("Item 2");
                cells["B3"].PutValue(150);
                cells["A4"].PutValue("Item 3");
                cells["B4"].PutValue(200);
                cells["A5"].PutValue("Item 4");
                cells["B5"].PutValue(250);

                // Insert a new column C for the running total
                cells.InsertColumn(2); // zero‑based index: 2 => column C
                cells["C1"].PutValue("Running Total");

                // Set running total formula for each data row
                // Formula: =SUM($B$2:B2) for row 2, =SUM($B$2:B3) for row 3, etc.
                int startDataRow = 1; // zero‑based index for row 2 (first data row)
                int amountColIndex = 1; // column B
                int totalColIndex = 2; // column C
                int lastDataRow = 5; // row 6 in zero‑based (including header)

                for (int row = startDataRow; row <= lastDataRow; row++)
                {
                    // Build absolute reference to the first amount cell and relative reference to current amount cell
                    string firstAmountCell = CellsHelper.CellIndexToName(startDataRow, amountColIndex); // e.g., B2
                    string currentAmountCell = CellsHelper.CellIndexToName(row, amountColIndex); // e.g., B3, B4, ...

                    // Construct SUM formula
                    string formula = $"=SUM(${firstAmountCell}:{currentAmountCell})";

                    // Apply formula to the running total cell
                    cells[row, totalColIndex].Formula = formula;
                }

                // Ensure output directory exists
                string outputPath = "RunningTotalExample.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
