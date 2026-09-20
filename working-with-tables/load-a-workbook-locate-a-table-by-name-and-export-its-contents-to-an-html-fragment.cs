// Title: Export a specific Excel ListObject to an HTML fragment with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook, locates a ListObject by its name, copies its data range into a new workbook, and returns the result as an HTML string using Aspose.Cells. | Create a method that extracts a named table from any worksheet, preserves the original column widths, and saves it as an HTML fragment with embedded base64 images via Aspose.Cells HtmlSaveOptions.
// Common Searches: Aspose.Cells C# export named table to HTML fragment | How to convert a ListObject to HTML using Aspose.Cells | Extract Excel table range and get HTML string with base64 images in .NET | Save only the worksheet that contains a specific table as HTML with Aspose.Cells
// Tags: export ListObject to HTML Aspose.Cells | named Excel table HTML conversion C# | HtmlSaveOptions ExportActiveWorksheetOnly Aspose.Cells | copy table range to new workbook Aspose.Cells | embed images base64 in HTML export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads a workbook, searches all worksheets for a ListObject named "MyTable", copies its data range and column widths into a new workbook, and then uses HtmlSaveOptions with ExportActiveWorksheetOnly and base64 image embedding to produce an HTML fragment representing the table.
class ExportTableToHtml
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "SourceWorkbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Name of the table (ListObject) to export
            string tableName = "MyTable";

            // Load the workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Locate the table by name across all worksheets
            ListObject? targetTable = null;
            Worksheet? tableWorksheet = null;

            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Try to get the table by name; returns null if not found
                ListObject? lo = ws.ListObjects[tableName];
                if (lo != null)
                {
                    targetTable = lo;
                    tableWorksheet = ws;
                    break;
                }
            }

            if (targetTable == null || tableWorksheet == null)
            {
                Console.WriteLine($"Table \"{tableName}\" not found in the workbook.");
                return;
            }

            // Create a new workbook to hold only the table data
            Workbook htmlWorkbook = new Workbook();
            Worksheet htmlSheet = htmlWorkbook.Worksheets[0];
            // Use the table's display name as the worksheet name
            htmlSheet.Name = targetTable.DisplayName;

            // Determine the range of the table (including headers)
            int startRow = targetTable.DataRange.FirstRow;
            int startColumn = targetTable.DataRange.FirstColumn;
            int rowCount = targetTable.DataRange.RowCount;
            int columnCount = targetTable.DataRange.ColumnCount;

            // Copy rows and columns from source to destination
            htmlSheet.Cells.CopyRows(tableWorksheet.Cells, startRow, 0, rowCount);
            htmlSheet.Cells.CopyColumns(tableWorksheet.Cells, startColumn, 0, columnCount);

            // Adjust column widths to match the source
            for (int col = 0; col < columnCount; col++)
            {
                int sourceColIndex = startColumn + col;
                int widthPixel = (int)tableWorksheet.Cells.GetColumnWidthPixel(sourceColIndex);
                htmlSheet.Cells.SetColumnWidthPixel(col, widthPixel);
            }

            // Prepare HTML save options to export only the prepared worksheet
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true,
                ExportImagesAsBase64 = true
                // The output will be a full HTML document containing only the active worksheet.
            };

            // Save the HTML to a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                htmlWorkbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0;

                // Convert the stream to a string (HTML)
                using (StreamReader reader = new StreamReader(htmlStream))
                {
                    string htmlContent = reader.ReadToEnd();

                    Console.WriteLine("HTML output of the table:");
                    Console.WriteLine(htmlContent);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
