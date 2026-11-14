// Title: Add a clickable hyperlink to a specific cell in an Aspose.Cells ListObject table using C#
// Common Searches: C# Aspose.Cells put hyperlink text into a ListObject cell | How to add a URL to a column of an Excel table using Aspose.Cells | Insert clickable link in Aspose.Cells table with PutCellValue and Hyperlinks | Aspose.Cells .NET create hyperlink inside a table row | Set cell hyperlink in Excel ListObject using Aspose.Cells API
// Tags: ListObject.PutCellValue set hyperlink text | Aspose.Cells add hyperlink to table cell | HyperlinkCollection.Add single cell Excel | C# insert URL into ListObject column | Excel table hyperlink insertion Aspose.Cells | Clickable link generation with Aspose.Cells worksheet

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

namespace Example
{
    // The example creates a workbook, defines a ListObject covering A1:C4, uses ListObject.PutCellValue to place a URL string in a specific table cell, adds a clickable hyperlink to that cell via HyperlinkCollection.Add, and saves the file as Output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data to define the table range (A1:C4)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Link");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Item1");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Item2");
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Item3");

                // Create a ListObject (table) covering the range A1:C4 (hasHeaders = true)
                int firstRow = 0;          // zero‑based index for row 1
                int firstColumn = 0;       // zero‑based index for column A
                int totalRows = 4;
                int totalColumns = 3;
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";

                // Insert a hyperlink string into the third column of the second row of the table
                int rowIndex = 1;          // second row inside the table (zero‑based)
                int columnIndex = 2;       // third column inside the table (zero‑based)
                string hyperlink = "http://example.com";

                // Set the display text in the cell
                table.PutCellValue(rowIndex, columnIndex, hyperlink);

                // Add an actual hyperlink to the cell so Excel treats it as clickable
                int targetRow = table.StartRow + rowIndex;
                int targetColumn = table.StartColumn + columnIndex;
                // HyperlinkCollection.Add overload requires range size; use 1x1 for a single cell
                sheet.Hyperlinks.Add(targetRow, targetColumn, 1, 1, hyperlink);

                // Save the workbook
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
