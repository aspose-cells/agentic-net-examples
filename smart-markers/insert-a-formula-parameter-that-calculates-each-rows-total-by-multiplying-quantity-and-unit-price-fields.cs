// Title: C# – Add a row‑wise total formula to an Aspose.Cells ListObject (Quantity × UnitPrice)
// Description: Creates a workbook, defines a table with Quantity, UnitPrice and Total columns, applies the structured reference formula "[@Quantity]*[@UnitPrice]" to every data row of the Total column via ListObject.PutCellFormula, recalculates the sheet, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | ListObject | structured reference | Excel table formula | multiply columns | row total calculation | Workbook.CalculateFormula | export to Excel | table column formula
// Common Searches: Aspose.Cells add formula to ListObject column | C# set structured reference formula in Excel table | calculate total column in Aspose.Cells workbook | how to use PutCellFormula with ListObject | multiply two columns in Aspose.Cells table
// Developer Intent: Insert a formula that multiplies Quantity by UnitPrice for each data row in the Total column of a ListObject.
// Use Cases: Automatically compute line‑item totals in an invoice generated with Aspose.Cells. | Build a sales report where totals update instantly when Quantity or UnitPrice values change. | Export database records to Excel with pre‑filled calculation formulas for downstream analysis.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, adds a ListObject, and sets a structured formula "[@Quantity]*[@UnitPrice]" for the Total column. | Show how to recalculate all formulas after inserting them into an Aspose.Cells table and then save the workbook. | Explain the use of structured references like [@ColumnName] in Aspose.Cells formulas for table‑based calculations.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsFormulaDemo
{
    // Creates a workbook, defines a table with Quantity, UnitPrice and Total columns, applies the structured reference formula "[@Quantity]*[@UnitPrice]" to every data row of the Total column via ListObject.PutCellFormula, recalculates the sheet, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add headers for Quantity, UnitPrice and Total
                sheet.Cells["A1"].PutValue("Quantity");
                sheet.Cells["B1"].PutValue("UnitPrice");
                sheet.Cells["C1"].PutValue("Total");

                // Populate some sample data (Quantity, UnitPrice)
                sheet.Cells["A2"].PutValue(5);
                sheet.Cells["B2"].PutValue(12.5);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["B3"].PutValue(7.8);
                sheet.Cells["A4"].PutValue(10);
                sheet.Cells["B4"].PutValue(4.2);

                // Create a ListObject (table) that covers the data range including the header
                // Parameters: first row, first column, last row, last column, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Apply a formula to the "Total" column for each data row.
                // Structured reference "[@Quantity]" and "[@UnitPrice]" refer to the current row's cells.
                string totalFormula = "=[@Quantity]*[@UnitPrice]";

                // Calculate number of data rows (exclude header row)
                int dataRows = table.DataRange.RowCount - 1;

                // Data rows start at offset 1 (skip header), column offset 2 (third column = Total)
                for (int i = 0; i < dataRows; i++)
                {
                    table.PutCellFormula(i + 1, 2, totalFormula);
                }

                // Optionally calculate formulas so the workbook shows results immediately
                workbook.CalculateFormula();

                // Save the workbook
                workbook.Save("TableWithTotalFormula.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
