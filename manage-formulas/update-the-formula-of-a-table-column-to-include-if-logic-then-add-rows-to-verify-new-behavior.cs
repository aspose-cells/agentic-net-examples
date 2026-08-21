// Title: C# – Update a ListObject column formula with IF logic and add rows – Aspose.Cells example
// Description: Creates a workbook with a table (ID, Value, Status), applies a conditional IF formula to the Status column, inserts new rows, recalculates formulas, prints the results, and saves the file as UpdatedTableFormula.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# table formula | ListObject IF formula .NET | update column formula Aspose.Cells | add rows to ListObject | recalculate formulas Aspose.Cells | custom calculated formula | Excel table conditional formatting code | Aspose.Cells example GitHub
// Common Searches: how to set IF formula on a table column with Aspose.Cells | add rows to ListObject and recalculate formulas C# | update ListObject column formula programmatically | Aspose.Cells example for conditional column values | C# code to verify table formulas after inserting rows
// Developer Intent: Programmatically change a table column's formula to include conditional logic and ensure the new formula is applied to rows added later.
// Use Cases: Generate a dynamic Status column that labels values as High or Low based on a threshold. | Automatically compute column values for rows appended after the initial table creation. | Produce an Excel workbook with up‑to‑date calculations without manual formula entry.
// AI Prompts: Write C# code that sets an IF formula on a ListObject column, adds new rows, triggers recalculation, and saves the workbook with Aspose.Cells. | Explain the SetCustomCalculatedFormula parameters and how they control formula propagation in Aspose.Cells tables. | Show how to iterate over a table's DataRange to confirm that newly added rows have correctly evaluated formulas.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook with a table (ID, Value, Status), applies a conditional IF formula to the Status column, inserts new rows, recalculates formulas, prints the results, and saves the file as UpdatedTableFormula.xlsx using Aspose.Cells for .NET.
    class UpdateTableColumnFormula
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // ----- Create sample data -----
                ws.Cells["A1"].PutValue("ID");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["C1"].PutValue("Status"); // Column that will hold the IF result

                ws.Cells["A2"].PutValue(1);
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue(2);
                ws.Cells["B3"].PutValue(25);
                ws.Cells["A4"].PutValue(3);
                ws.Cells["B4"].PutValue(5);

                // ----- Create a table (ListObject) covering the data range -----
                int tableIndex = ws.ListObjects.Add("A1", "C4", true);
                ListObject table = ws.ListObjects[tableIndex];
                table.DisplayName = "DataTable";

                // ----- Update the formula of the "Status" column to include IF logic -----
                // Formula: =IF([@Value]>15,"High","Low")
                ListColumn statusColumn = table.ListColumns[2]; // third column (zero‑based index)
                statusColumn.SetCustomCalculatedFormula("=IF([@Value]>15,\"High\",\"Low\")", false, false);

                // ----- Add new rows to verify that the formula is applied automatically -----
                // Row offsets are zero‑based within the table's data rows (excluding the header)
                table.PutCellValue(3, 0, 4); // ID = 4
                table.PutCellValue(3, 1, 30); // Value = 30

                table.PutCellValue(4, 0, 5); // ID = 5
                table.PutCellValue(4, 1, 8); // Value = 8

                // Recalculate formulas so that the new rows get their Status values
                wb.CalculateFormula();

                // ----- Output the table content to the console for verification -----
                Console.WriteLine("ID\tValue\tStatus");
                AsposeRange dataRange = table.DataRange; // Range that contains the data rows
                int startRow = dataRange.FirstRow;
                int endRow = dataRange.FirstRow + dataRange.RowCount - 1;
                for (int row = startRow; row <= endRow; row++)
                {
                    int id = ws.Cells[row, 0].IntValue;
                    int value = ws.Cells[row, 1].IntValue;
                    string status = ws.Cells[row, 2].StringValue;
                    Console.WriteLine($"{id}\t{value}\t{status}");
                }

                // ----- Save the workbook -----
                string outputPath = "UpdatedTableFormula.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
