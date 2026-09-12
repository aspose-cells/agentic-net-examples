// Title: Export a named Excel table to an Excel 2003 XML file using Aspose.Cells for .NET
// AI Prompts: Write C# that opens an .xlsx workbook, locates a ListObject by its name, copies the entire table (including headers and styles) to a new workbook, and saves it as Excel 2003 XML with Aspose.Cells. | Create a reusable method that checks if the source file and the specified table exist, then writes the table data to an XML document while preserving cell formatting using Aspose.Cells SaveFormat.Xml. | Generate a code snippet that extracts a table range from a worksheet, transfers values and styles to a temporary worksheet, and writes the result to an XML file compatible with Excel 2003.
// Common Searches: Aspose.Cells C# convert a named ListObject into an XML document | How to save an Excel table as Excel 2003 XML using .NET | C# example for copying a named table with formatting and exporting to XML | Validate table existence before exporting to XML with Aspose.Cells | Export Excel table to XML format preserving styles Aspose.Cells
// Tags: export ListObject to Excel 2003 XML | copy table range with formatting Aspose.Cells | save workbook as XML using SaveFormat.Xml | validate source file and table existence .NET | named table extraction from .xlsx | preserve cell styles during XML export

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// Loads 'source.xlsx', finds the ListObject named 'MyTable', copies its cells and styles to a new workbook, and saves the result as 'ExportedTable.xml' in Excel 2003 XML format, with checks for missing source file or table.
class ExportTableToXml
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "ExportedTable.xml";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook that contains the table.
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Retrieve the table (ListObject) by its name. Replace "MyTable" with the actual table name.
            ListObject table = srcSheet.ListObjects["MyTable"];
            if (table == null)
            {
                Console.WriteLine("Table 'MyTable' not found in the worksheet.");
                return;
            }

            // Determine the range that covers the entire table (including header).
            int startRow = table.StartRow;                                   // First row of the table (header)
            int rowCount = table.EndRow - table.StartRow + 1;                // Total rows including header
            int startColumn = table.StartColumn;                             // First column of the table
            int columnCount = table.EndColumn - table.StartColumn + 1;       // Total columns

            // Create a new workbook that will hold only the exported table.
            Workbook exportWorkbook = new Workbook();
            Worksheet exportSheet = exportWorkbook.Worksheets[0];
            exportSheet.Name = "ExportedTable";

            // Copy the table's cells to the new worksheet.
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < columnCount; c++)
                {
                    Cell srcCell = srcSheet.Cells[startRow + r, startColumn + c];
                    Cell destCell = exportSheet.Cells[r, c];
                    destCell.PutValue(srcCell.Value);
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Save the new workbook using the Excel 2003 XML format.
            exportWorkbook.Save(outputPath, SaveFormat.Xml);
            Console.WriteLine($"Table exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
