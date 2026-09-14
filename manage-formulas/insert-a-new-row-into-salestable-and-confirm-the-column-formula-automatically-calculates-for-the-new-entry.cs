// Title: Insert a new row into an Excel ListObject and automatically recalculate its formula column using Aspose.Cells for .NET
// AI Prompts: Add a data row to the 'SalesTable' ListObject, set Product, Quantity, and UnitPrice values, then invoke Workbook.CalculateFormula so the Total column is computed for the new entry. | Programmatically insert a row after the last record of an Excel table, let the table expand automatically, and fetch the calculated result of the formula column with Aspose.Cells in C#.
// Common Searches: C# Aspose.Cells add row to ListObject and keep calculated columns updated | How to expand an Excel table and recalculate formulas after inserting a row with Aspose.Cells | Retrieve formula result for newly inserted row in Excel table using Aspose.Cells .NET | Insert new record into SalesTable and get Total column value programmatically
// Tags: insert row ListObject Aspose.Cells | auto expand Excel table C# | recalculate formulas workbook.CalculateFormula | retrieve calculated column value Aspose.Cells | populate SalesTable new record

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an existing workbook, locates the ListObject named 'SalesTable', inserts a new data row after the current rows, assigns values for Product, Quantity, and UnitPrice, triggers Workbook.CalculateFormula to compute the Total column for the new row, prints the calculated result, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "SalesData.xlsx";
            const string outputPath = "SalesData_Updated.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook that contains the SalesTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (assumed to hold the table)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the ListObject (table) named "SalesTable"
            ListObject salesTable = sheet.ListObjects["SalesTable"];
            if (salesTable == null)
            {
                Console.WriteLine("Table 'SalesTable' not found in the worksheet.");
                return;
            }

            // Determine the row index where the new row will be inserted (after the last data row)
            int insertRowIndex = salesTable.DataRange.FirstRow + salesTable.DataRange.RowCount;

            // Insert a single new row into the worksheet at the calculated position.
            // The table will automatically expand to include the new row.
            sheet.Cells.InsertRows(insertRowIndex, 1);

            // Set values for the newly inserted row (except the formula column)
            // Example column order: Product (0), Quantity (1), UnitPrice (2), Total (3 - formula)
            int firstDataCol = salesTable.DataRange.FirstColumn;
            sheet.Cells[insertRowIndex, firstDataCol + 0].PutValue("NewProduct");
            sheet.Cells[insertRowIndex, firstDataCol + 1].PutValue(10);          // Quantity
            sheet.Cells[insertRowIndex, firstDataCol + 2].PutValue(15.75);      // UnitPrice

            // Recalculate all formulas so that the formula column (e.g., Total) updates automatically
            workbook.CalculateFormula();

            // Verify that the formula column has been calculated for the new row
            int formulaColIndex = firstDataCol + salesTable.DataRange.ColumnCount - 1; // last column in the table
            object calculatedValue = sheet.Cells[insertRowIndex, formulaColIndex].Value;
            Console.WriteLine("Calculated formula result for the new row: " + calculatedValue);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
