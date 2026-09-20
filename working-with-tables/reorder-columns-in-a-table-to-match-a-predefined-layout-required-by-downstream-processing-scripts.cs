// Title: Reorder Excel ListObject columns to a specific header sequence using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx workbook, extract the first ListObject into a DataTable, rearrange its columns according to a predefined header array, and write the reordered data back with Aspose.Cells in C#. | Clear the original table range, insert headers in a new order, populate rows from the reordered DataTable, and resize the ListObject to fit the updated layout using Aspose.Cells. | Create a C# routine that checks for required column names, changes the column order of an Excel table, and saves the modified workbook to a new file with Aspose.Cells.
// Common Searches: C# Aspose.Cells reorder columns in an Excel table based on custom header list | How to change column order of a ListObject in an .xlsx file using Aspose.Cells .NET | Resize Excel ListObject after modifying its data with Aspose.Cells | Move Excel table columns to match a predefined layout programmatically in C# | Validate column names before reordering an Excel table using Aspose.Cells
// Tags: Aspose.Cells rearrange ListObject column order | Excel table column ordering C# | DataTable column mapping Aspose.Cells | Resize ListObject after data rewrite | Validate Excel table headers with Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an .xlsx workbook, accesses the first worksheet's ListObject, extracts its data into a DataTable, rearranges columns to match a predefined header sequence (ID, Name, Date, Amount), clears the original range, writes the reordered headers and rows back, resizes the table to the new layout, and saves the workbook.
class ReorderTableColumns
{
    static void Main()
    {
        // Define file paths
        string inputPath = @"C:\Data\InputWorkbook.xlsx";
        string outputPath = @"C:\Data\OutputWorkbook.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }

            // Use the first table as the target
            ListObject table = sheet.ListObjects[0];

            // Desired column order (must match column names in the table header)
            string[] desiredOrder = { "ID", "Name", "Date", "Amount" };

            // Determine table boundaries
            int firstRow = table.StartRow;          // Header row index
            int firstCol = table.StartColumn;

            // DataRange includes only data rows (no header)
            int dataRows = table.DataRange.RowCount;
            int totalRows = dataRows + 1;           // Include header row
            int totalCols = table.DataRange.ColumnCount;

            // Load the table data into a DataTable
            DataTable sourceTable = new DataTable();

            // Add columns using header values
            for (int c = 0; c < totalCols; c++)
            {
                string header = sheet.Cells[firstRow, firstCol + c].StringValue;
                sourceTable.Columns.Add(header);
            }

            // Add rows (skip header)
            for (int r = 1; r < totalRows; r++)
            {
                DataRow row = sourceTable.NewRow();
                for (int c = 0; c < totalCols; c++)
                {
                    row[c] = sheet.Cells[firstRow + r, firstCol + c].Value;
                }
                sourceTable.Rows.Add(row);
            }

            // Create a new DataTable with columns in the desired order
            DataTable reorderedTable = new DataTable();

            foreach (string colName in desiredOrder)
            {
                if (!sourceTable.Columns.Contains(colName))
                {
                    Console.WriteLine($"Column '{colName}' not found in source table.");
                    return;
                }
                reorderedTable.Columns.Add(colName, sourceTable.Columns[colName].DataType);
            }

            // Populate rows according to the new column order
            foreach (DataRow srcRow in sourceTable.Rows)
            {
                DataRow newRow = reorderedTable.NewRow();
                foreach (string colName in desiredOrder)
                {
                    newRow[colName] = srcRow[colName];
                }
                reorderedTable.Rows.Add(newRow);
            }

            // Clear the original table range (including header)
            sheet.Cells.CreateRange(firstRow, firstCol, totalRows, totalCols).ClearContents();

            // Write the reordered data back to the worksheet (headers first)
            for (int c = 0; c < desiredOrder.Length; c++)
            {
                sheet.Cells[firstRow, firstCol + c].PutValue(desiredOrder[c]);
            }

            // Then write data rows
            for (int r = 0; r < reorderedTable.Rows.Count; r++)
            {
                for (int c = 0; c < desiredOrder.Length; c++)
                {
                    sheet.Cells[firstRow + 1 + r, firstCol + c].PutValue(reorderedTable.Rows[r][c]);
                }
            }

            // Resize the table to match the new layout (hasHeaders = true)
            table.Resize(firstRow, firstCol, totalRows, totalCols, true);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine("Columns reordered and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
