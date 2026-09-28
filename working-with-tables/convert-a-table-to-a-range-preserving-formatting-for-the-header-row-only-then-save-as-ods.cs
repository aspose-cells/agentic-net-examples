// Title: Convert an Excel ListObject to a plain range, keep header formatting, and save as ODS using Aspose.Cells for .NET
// AI Prompts: Delete the table object from a worksheet while preserving the header row style, then export the workbook to ODS. | Clear formatting of all data rows in a ListObject, remove the table definition, and save the file as an ODS document with Aspose.Cells.
// Common Searches: Aspose.Cells how to delete a table but retain header formatting in C# | convert Excel table to range and export to ODS using .NET | remove ListObject from worksheet and keep header style Aspose.Cells | save workbook as ODS after clearing table data row styles | C# code to transform Excel ListObject into plain cells and output ODS
// Tags: listobject deletion keep header formatting Aspose.Cells | plain range conversion from table C# | ods export Aspose.Cells .NET | data rows style reset Aspose.Cells | excel to ods conversion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads or creates an Excel workbook, accesses the first ListObject, clears styles from all data rows while preserving the header row's formatting, removes the table definition, and saves the resulting worksheet as an ODS file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Load existing workbook or create a new one if the file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a simple workbook with a sample table to work with.
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["A2"].PutValue("Data1");
                ws.Cells["B2"].PutValue("Data2");

                // Add a table covering the sample data.
                int firstRow = 0, firstColumn = 0, totalRows = 2, totalColumns = 2;
                int tableIndex = ws.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows, firstColumn + totalColumns - 1, true);
                ListObject table = ws.ListObjects[tableIndex];
                table.DisplayName = "SampleTable";
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first table (ListObject) on the worksheet.
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            ListObject tableObj = sheet.ListObjects[0];

            // Determine the bounds of the table (including the header row).
            int startRow = tableObj.StartRow;          // Header row index
            int startColumn = tableObj.StartColumn;    // First column index
            int endRow = tableObj.EndRow;
            int endColumn = tableObj.EndColumn;
            int rowCount = endRow - startRow + 1;
            int columnCount = endColumn - startColumn + 1;

            // Create a default (empty) style.
            Style defaultStyle = workbook.CreateStyle();

            // Clear formatting for all data rows while preserving the header row.
            for (int r = startRow + 1; r <= endRow; r++) // skip header row
            {
                for (int c = startColumn; c <= endColumn; c++)
                {
                    sheet.Cells[r, c].SetStyle(defaultStyle);
                }
            }

            // Remove the table definition, leaving a plain range with the same cells.
            int tableIndexToRemove = sheet.ListObjects.IndexOf(tableObj);
            if (tableIndexToRemove >= 0)
            {
                sheet.ListObjects.RemoveAt(tableIndexToRemove);
            }

            // Save the workbook as an ODS file.
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
