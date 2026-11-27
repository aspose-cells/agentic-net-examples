// Title: Import a DataTable into an Excel worksheet at row 0 to overwrite existing content with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import DataTable starting at first row to replace existing sheet data | how to set firstRow offset to zero when importing a DataTable with Aspose.Cells | overwrite existing rows in an Excel worksheet using Aspose.Cells import dataset | C# write DataTable headers and rows to Excel sheet without shifting existing cells Aspose.Cells | replace worksheet content with new DataTable using Aspose.Cells for .NET
// Tags: import DataTable to Excel worksheet Aspose.Cells | firstRow offset zero Aspose.Cells | overwrite worksheet rows Aspose.Cells | write DataTable headers Aspose.Cells | save workbook as Xlsx Aspose.Cells | manual DataTable import C# Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsGridWebImportDemo
{
    // The example creates a DataTable, opens or adds a worksheet in a new Workbook, sets the start row and column to zero, writes column headers and row values directly to the cells, and saves the file as an Xlsx. Setting the firstRow offset to zero ensures the imported data overwrites any existing content on the sheet.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Prepare a DataTable that will act as the source dataset
            // ------------------------------------------------------------
            DataTable sourceTable = new DataTable("People");
            sourceTable.Columns.Add("Name", typeof(string));
            sourceTable.Columns.Add("Age", typeof(int));

            sourceTable.Rows.Add("Alice", 30);
            sourceTable.Rows.Add("Bob", 25);
            sourceTable.Rows.Add("Charlie", 35);

            // ------------------------------------------------------------
            // 2. Create a Workbook and a target worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();

            string targetSheetName = "ImportedData";
            Worksheet sheet;

            // Try to get an existing sheet; otherwise create a new one
            try
            {
                sheet = workbook.Worksheets[targetSheetName];
            }
            catch
            {
                int index = workbook.Worksheets.Add();
                sheet = workbook.Worksheets[index];
                sheet.Name = targetSheetName;
            }

            // ------------------------------------------------------------
            // 3. Import the DataTable into the worksheet at row 0, column 0
            // ------------------------------------------------------------
            int startRow = 0;      // firstRow offset = 0 (overwrite existing data)
            int startColumn = 0;   // firstColumn offset = 0
            bool importColumnNames = true; // include header row

            try
            {
                // Manual import (ImportDataTable may not be available in some versions)
                int rowOffset = startRow;
                int colOffset = startColumn;

                if (importColumnNames)
                {
                    // Write column headers
                    for (int c = 0; c < sourceTable.Columns.Count; c++)
                    {
                        sheet.Cells[rowOffset, colOffset + c].PutValue(sourceTable.Columns[c].ColumnName);
                    }
                    rowOffset++; // Move to next row for data
                }

                // Write data rows
                for (int r = 0; r < sourceTable.Rows.Count; r++)
                {
                    for (int c = 0; c < sourceTable.Columns.Count; c++)
                    {
                        sheet.Cells[rowOffset + r, colOffset + c].PutValue(sourceTable.Rows[r][c]);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Import failed: " + ex.Message);
                return;
            }

            // ------------------------------------------------------------
            // 4. Verify that the data was imported at the expected location
            // ------------------------------------------------------------
            Console.WriteLine("Imported data (including header):");
            int totalRows = sourceTable.Rows.Count + (importColumnNames ? 1 : 0);
            for (int r = 0; r < totalRows; r++)
            {
                string name = sheet.Cells[r, 0].StringValue;
                string age = sheet.Cells[r, 1].StringValue;
                Console.WriteLine($"Row {r}: Name = {name}, Age = {age}");
            }

            // ------------------------------------------------------------
            // 5. Save the workbook to an Excel file
            // ------------------------------------------------------------
            string outputPath = "ImportedData.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Saving to Excel failed: " + ex.Message);
            }
        }
    }
}
