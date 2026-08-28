// Title: Generate an Excel file with Aspose.Cells smart markers that calculate a running‑total column using formulas referencing previous rows in C#
// AI Prompts: Write C# code that creates a workbook, adds smart markers for Name and Amount, processes a List<Record> data source, and then inserts a cumulative‑sum formula into the column that stores the running total. | Demonstrate using Aspose.Cells Table.PutCellFormula after smart marker processing to set formulas such as =B2 and =B3+C2 for each row, recalculate the workbook, and save the file.
// Common Searches: how to compute a cumulative sum column with Aspose.Cells smart markers in C# | c# Aspose.Cells set formula referencing previous row in a ListObject table | example of using Table.PutCellFormula for running total after smart marker expansion | generate Excel report with smart markers and calculate totals per row using Aspose.Cells
// Tags: Aspose.Cells smart marker total column | C# Table.PutCellFormula previous row reference | Excel cumulative column using smart markers | dynamic ListObject with formulas Aspose.Cells | apply formulas after smart marker processing C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSmartMarkerRunningTotal
{
    // Simple data class for the smart marker source
    // The example creates a workbook, defines smart markers for Name and Amount, processes a List<Record> data source into a ListObject table, then programmatically adds running‑total formulas in column C that reference the current Amount cell and the previous row's total, recalculates all formulas, and saves the result as an Excel file.
    public class Record
    {
        public string Name { get; set; }
        public double Amount { get; set; }

        public Record(string name, double amount)
        {
            Name = name;
            Amount = amount;
        }
    }

    public class RunningTotalDemo
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook (lifecycle rule) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Set up the template with smart markers ----------
                // Header row
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Amount");
                cells["C1"].PutValue("RunningTotal");

                // Data row with smart markers (will be repeated)
                cells["A2"].PutValue("&=$Name");
                cells["B2"].PutValue("&=$Amount");
                // Column C (RunningTotal) is left empty; we'll fill it with formulas later

                // ---------- Create a table that covers the template range ----------
                // The table will automatically expand when the smart markers are processed
                int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // ---------- Prepare data source ----------
                List<Record> data = new List<Record>
                {
                    new Record("Item A", 100),
                    new Record("Item B", 150),
                    new Record("Item C", 200),
                    new Record("Item D", 250),
                    new Record("Item E", 300)
                };

                // ---------- Process smart markers ----------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Data", data);
                designer.Process();

                // ---------- Insert running‑total formulas ----------
                // The table now contains the data rows. We set a formula for each row in column C.
                // Row offset 0 = header, so data rows start at offset 1.
                for (int rowOffset = 1; rowOffset <= data.Count; rowOffset++)
                {
                    // Convert table row offset to worksheet row number (1‑based for Excel)
                    int sheetRow = table.StartRow + rowOffset + 1; // +1 because Excel rows start at 1

                    string formula;
                    if (rowOffset == 1)
                    {
                        // First row: running total equals the amount of the first row
                        formula = $"=B{sheetRow}";
                    }
                    else
                    {
                        // Subsequent rows: current amount + previous running total
                        int prevSheetRow = sheetRow - 1;
                        formula = $"=B{sheetRow}+C{prevSheetRow}";
                    }

                    // Put the formula into column C (offset 2) of the current table row
                    table.PutCellFormula(rowOffset, 2, formula);
                }

                // ---------- Calculate all formulas ----------
                workbook.CalculateFormula();

                // ---------- Save the workbook (lifecycle rule) ----------
                workbook.Save("RunningTotalSmartMarker.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
