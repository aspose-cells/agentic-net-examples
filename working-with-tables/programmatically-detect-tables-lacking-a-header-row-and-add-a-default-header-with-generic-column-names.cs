// Title: C# – Detect Tables Without Headers and Add Generic Column Names Using AspNet.Cells for .NET
// Description: This example creates a workbook, adds a ListObject (Excel table) without a header row, scans every worksheet, inserts a new top row for each header‑less table, fills it with generic names (Column1, Column2, …), calls ListObject.UpdateColumnName to sync the table definition, and saves the result as an .xlsx file.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET | detect missing table header | add default header row | generic column names Excel | ListObject UpdateColumnName | programmatically add table header | Excel table without headers | C# Excel automation | Aspose.Cells sample code
// Common Searches: how to add a header row to an Aspose.Cells ListObject | detect tables without headers in a workbook using Aspose.Cells | insert generic column names into Excel tables C# | Aspose.Cells update column names after inserting header | C# code to add default headers to all tables in Excel file
// Developer Intent: Automatically insert a default header row with generic column names into any ListObject that lacks one.
// Use Cases: Ensure every table in a generated report has a header before publishing. | Standardize imported spreadsheets that miss header rows by adding Column1, Column2, … automatically. | Prepare workbooks for downstream analytics that require defined column names for each table.
// AI Prompts: Write C# code with Aspose.Cells that scans a workbook, finds ListObjects without headers, inserts a header row named Column1, Column2, etc., and updates the table definition. | Explain why ListObject.UpdateColumnName must be called after adding a header row in Aspose.Cells. | Create a reusable method that accepts a Workbook object, adds generic headers to any header‑less table, and returns the modified workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a ListObject (Excel table) without a header row, scans every worksheet, inserts a new top row for each header‑less table, fills it with generic names (Column1, Column2, …), calls ListObject.UpdateColumnName to sync the table definition, and saves the result as an .xlsx file.
    class DetectAndAddTableHeaders
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample data without a header row (table starts at A1)
                worksheet.Cells["A1"].PutValue("Apple");
                worksheet.Cells["B1"].PutValue(10);
                worksheet.Cells["A2"].PutValue("Orange");
                worksheet.Cells["B2"].PutValue(15);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(8);

                // Add a ListObject (table) without headers (showHeaders = false)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, false);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "FruitTable";

                // Iterate through all worksheets and their tables
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (ListObject lo in ws.ListObjects)
                    {
                        // Determine the range of the table
                        int startRow = lo.StartRow;
                        int startColumn = lo.StartColumn;
                        int columnCount = lo.EndColumn - lo.StartColumn + 1;

                        // Insert a new row at the beginning of the table range
                        ws.Cells.InsertRow(startRow);

                        // Populate generic column names: Column1, Column2, ...
                        for (int c = 0; c < columnCount; c++)
                        {
                            ws.Cells[startRow, startColumn + c].PutValue($"Column{c + 1}");
                        }

                        // Update the ListObject's column names to match the new header cells
                        lo.UpdateColumnName();
                    }
                }

                // Define output file path
                string outputPath = "TablesWithHeaders.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
