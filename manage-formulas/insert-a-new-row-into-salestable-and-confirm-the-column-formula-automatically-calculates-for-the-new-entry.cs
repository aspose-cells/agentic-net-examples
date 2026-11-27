// Title: Insert a new row into a named Excel ListObject and automatically recalculate its column formulas with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a new row to a ListObject named SalesTable and keep column formulas intact | How to expand an Excel table range after inserting a row using Aspose.Cells in .NET | Force full formula recalculation for a table column after row insertion with Aspose.Cells C# | Populate values in a newly inserted ListObject row and retrieve the calculated Total column using Aspose.Cells
// Tags: insert blank row into ListObject Aspose.Cells C# | resize Excel table range after row insertion Aspose.Cells | recalculate table column formulas Aspose.Cells | populate new ListObject row values C# | retrieve calculated column value Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an existing workbook, finds the ListObject called "SalesTable", inserts a blank row after the last data row, expands the table to include the new row, writes sample values into the new cells, triggers a full formula calculation so that any table column formulas (e.g., a Total column) are evaluated for the added row, prints the calculated result, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook that contains the "SalesTable"
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the table (ListObject) named "SalesTable"
            ListObject salesTable = null;
            foreach (ListObject lo in worksheet.ListObjects)
            {
                // In Aspose.Cells the table name is accessed via DisplayName
                if (lo.DisplayName == "SalesTable")
                {
                    salesTable = lo;
                    break;
                }
            }

            if (salesTable == null)
            {
                Console.WriteLine("Table 'SalesTable' not found.");
                return;
            }

            // Determine the index where the new row should be inserted (after the last data row)
            int insertRowIndex = salesTable.DataRange.FirstRow + salesTable.DataRange.RowCount;

            // Insert a blank row in the worksheet at the calculated position
            worksheet.Cells.InsertRows(insertRowIndex, 1);

            // Expand the table range to include the newly inserted row
            // Assume the table has a header row (common case)
            int startRow = salesTable.DataRange.FirstRow - 1; // include header
            int startColumn = salesTable.DataRange.FirstColumn;
            int totalRows = salesTable.DataRange.RowCount + 2; // existing rows + header + new row
            int totalColumns = salesTable.DataRange.ColumnCount;

            // Resize the table; the last parameter indicates whether the table has headers
            salesTable.Resize(startRow, startColumn, totalRows, totalColumns, true);

            // Optionally populate some cells in the new row (example values)
            // Adjust column offsets according to your table structure
            worksheet.Cells[insertRowIndex, salesTable.DataRange.FirstColumn].PutValue("New Item");
            worksheet.Cells[insertRowIndex, salesTable.DataRange.FirstColumn + 1].PutValue(10);
            worksheet.Cells[insertRowIndex, salesTable.DataRange.FirstColumn + 2].PutValue(5.99);

            // Recalculate all formulas so that any column formulas in the table update automatically
            workbook.CalculateFormula();

            // Verify that a formula column (e.g., "Total") has been calculated for the new row
            // Assuming the formula column is the last column of the table
            int formulaColIndex = salesTable.DataRange.FirstColumn + salesTable.DataRange.ColumnCount - 1;
            object calculatedValue = worksheet.Cells[insertRowIndex, formulaColIndex].Value;
            Console.WriteLine($"Calculated value in formula column for new row: {calculatedValue}");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
