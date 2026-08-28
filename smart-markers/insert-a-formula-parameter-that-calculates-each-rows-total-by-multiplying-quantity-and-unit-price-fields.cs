// Title: Insert a structured reference formula into an Aspose.Cells ListObject to calculate each row's total (Quantity × UnitPrice) using C#
// AI Prompts: Set the Formula property of a ListColumn in a ListObject to a structured reference that multiplies the Quantity and UnitPrice columns for every row. | Enable a totals row on an Aspose.Cells table and configure the Total column to sum the calculated row totals.
// Common Searches: Aspose.Cells C# how to add a calculated column to a ListObject using structured references | C# set formula for table column in Aspose.Cells workbook | multiply two columns in an Aspose.Cells table and show sum in totals row | structured reference formula syntax for Aspose.Cells ListColumn | calculate row totals in an Excel table with Aspose.Cells .NET
// Tags: ListColumn.Formula structured reference Aspose.Cells | calculate row total in ListObject C# | totals row aggregation Aspose.Cells table | multiply Quantity UnitPrice Aspose.Cells | Excel table formula insertion .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsFormulaDemo
{
    // The example creates a workbook, defines a ListObject over a data range, assigns a structured reference formula ([@Quantity]*[@UnitPrice]) to the Total column, shows a totals row that sums the Total column, recalculates formulas, and saves the file as TableWithRowTotalFormula.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Populate sample data -----
            // Header row
            sheet.Cells["A1"].PutValue("Quantity");
            sheet.Cells["B1"].PutValue("UnitPrice");
            sheet.Cells["C1"].PutValue("Total");

            // Data rows
            sheet.Cells["A2"].PutValue(5);
            sheet.Cells["B2"].PutValue(12.5);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["B3"].PutValue(7.8);
            sheet.Cells["A4"].PutValue(10);
            sheet.Cells["B4"].PutValue(4.2);

            // ----- Create a table (ListObject) covering the data range -----
            // Table range: A1:C4 (including header)
            int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // ----- Insert formula for the Total column -----
            // The Total column is the third column in the table (index 2)
            // Use a structured reference that multiplies Quantity and UnitPrice for each row
            ListColumn totalColumn = table.ListColumns[2];
            totalColumn.Formula = "=[@Quantity]*[@UnitPrice]";

            // Optional: show totals row and calculate sum of totals
            table.ShowTotals = true;
            totalColumn.TotalsCalculation = TotalsCalculation.Sum;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("TableWithRowTotalFormula.xlsx");
        }
    }
}
