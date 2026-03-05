using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (header + 5 rows)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
            sheet.Cells[i - 1, 1].PutValue($"Student {i - 1}");       // Name
            sheet.Cells[i - 1, 2].PutValue(60 + i * 5);               // Score
        }

        // Create a named range that covers the data (A1:C6)
        AsposeRange dataRange = sheet.Cells.CreateRange("A1", "C6");
        dataRange.Name = "StudentData";

        // Add a ListObject (table) based on the same range
        int tableIndex = sheet.ListObjects.Add(
            dataRange.FirstRow,
            dataRange.FirstColumn,
            dataRange.RowCount,
            dataRange.ColumnCount,
            true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.DisplayName = "StudentsTable";
        table.ShowTotals = true;
        // Set totals calculation for the Score column (column index 2)
        table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum;

        // Export the table's data range to a DataTable
        DataTable dt = table.DataRange.ExportDataTable();

        // Output the exported DataTable to the console
        Console.WriteLine("Exported DataTable rows:");
        foreach (DataRow row in dt.Rows)
        {
            Console.WriteLine($"{row[0]}, {row[1]}, {row[2]}");
        }

        // Convert the table back to a normal range (removes table features)
        table.ConvertToRange();

        // Retrieve the previously named range
        AsposeRange retrievedRange = workbook.Worksheets.GetRangeByName("StudentData");
        if (retrievedRange != null)
        {
            int newColIndex = retrievedRange.FirstColumn + retrievedRange.ColumnCount; // D column

            // Add a new column header
            sheet.Cells[retrievedRange.FirstRow, newColIndex].PutValue("AdjustedScore");

            // Fill the new column with a formula (Score * 1.1)
            for (int r = 1; r < retrievedRange.RowCount; r++)
            {
                int rowIndex = retrievedRange.FirstRow + r;
                sheet.Cells[rowIndex, newColIndex].Formula = $"=C{rowIndex + 1}*1.1";
            }
        }

        // Save the workbook to an XLSX file
        workbook.Save("TablesAndRangesDemo.xlsx");
    }
}