// Title: Export a workbook with hidden rows and columns to PDF while preserving TableCssId only for visible tables using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, hides selected rows and columns, adds a ListObject, and saves the workbook as a PDF with Aspose.Cells. | Write C# logic that iterates through all ListObjects in a worksheet and omits assigning a TableCssId when the table contains any hidden rows or columns. | Modify the example to export the workbook to HTML instead of PDF, ensuring TableCssId is applied only to tables that have no hidden rows or columns.
// Common Searches: Aspose.Cells export workbook to PDF while ignoring tables that contain hidden rows | C# how to check for hidden rows or columns inside a ListObject using Aspose.Cells | TableCssId applied only to visible tables during HTML export Aspose.Cells .NET | Hide specific rows and columns before saving workbook as PDF with Aspose.Cells | Identify tables with hidden rows in Aspose.Cells and exclude them from CSS styling
// Tags: export workbook to PDF with hidden rows Aspose.Cells | detect hidden rows in ListObject C# | skip TableCssId for tables containing hidden columns Aspose.Cells | hide rows and columns before PDF export Aspose.Cells | visible table CSS identifier Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program creates a workbook, hides rows 4 and 7 and column B, adds a ListObject covering the data range, saves the workbook as a PDF, and then checks each table to determine if it contains hidden rows or columns, outputting messages that indicate whether the table should retain its TableCssId.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate some data (10 rows, 2 columns)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"R{i - 1}C1");
                sheet.Cells[$"B{i}"].PutValue($"R{i - 1}C2");
            }

            // Hide specific rows and columns (0‑based indices)
            sheet.Cells.Rows[3].IsHidden = true;   // Hide row 4
            sheet.Cells.Rows[6].IsHidden = true;   // Hide row 7
            sheet.Cells.Columns[1].IsHidden = true; // Hide column B

            // Add a table (ListObject) that spans the whole data range
            int firstRow = 0;   // A1
            int firstCol = 0;   // A1
            int totalRows = 10;
            int totalCols = 2;
            int tableIdx = sheet.ListObjects.Add(firstRow, firstCol,
                                                 firstRow + totalRows - 1,
                                                 firstCol + totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIdx];
            table.DisplayName = "VisibleTable";

            // Save the workbook as PDF
            string outputPath = Path.Combine(Environment.CurrentDirectory, "HiddenRowsColumns_Output.pdf");
            try
            {
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            }

            // Verify that tables containing hidden rows/columns are identified
            foreach (ListObject lo in sheet.ListObjects)
            {
                bool containsHidden = false;

                // Determine data range of the table
                Aspose.Cells.Range dataRange = lo.DataRange;
                int startRow = dataRange.FirstRow;
                int endRow = dataRange.FirstRow + dataRange.RowCount - 1;
                int startCol = dataRange.FirstColumn;
                int endCol = dataRange.FirstColumn + dataRange.ColumnCount - 1;

                // Check rows within the table range
                for (int r = startRow; r <= endRow; r++)
                {
                    if (sheet.Cells.Rows[r].IsHidden)
                    {
                        containsHidden = true;
                        break;
                    }
                }

                // If no hidden rows, check columns within the table range
                if (!containsHidden)
                {
                    for (int c = startCol; c <= endCol; c++)
                    {
                        if (sheet.Cells.Columns[c].IsHidden)
                        {
                            containsHidden = true;
                            break;
                        }
                    }
                }

                if (containsHidden)
                {
                    Console.WriteLine($"Table '{lo.DisplayName}' includes hidden rows/columns; it should be ignored for CSS ID.");
                }
                else
                {
                    Console.WriteLine($"Visible table '{lo.DisplayName}' retains its identifier.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
