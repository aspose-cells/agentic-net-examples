// Title: Create an Excel ListObject table with a visible comment using Aspose.Cells for .NET (C#)
// Common Searches: how to add a ListObject table to an Excel worksheet with Aspose.Cells in C# | Aspose.Cells C# add a visible comment to a specific cell | apply a built‑in table style and show first/last columns using Aspose.Cells | save a workbook containing a table and comment to a .xlsx file with Aspose.Cells | set comment author and visibility in Aspose.Cells for .NET
// Tags: Aspose.Cells create ListObject table C# | Aspose.Cells apply TableStyleMedium9 | Aspose.Cells add visible cell comment | Aspose.Cells show first and last column in table | Aspose.Cells save workbook as XLSX

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, adds a worksheet named "Data", fills cells A1:C4 with sample data, defines a ListObject table covering A1:C5, applies the TableStyleMedium9 style with first and last column highlighting, inserts a visible comment authored by "Automation" into cell A1, and saves the file as TableWithComment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Define the range that will become a table (list object)
            // Header row + 4 data rows = 5 rows, 3 columns (A:C)
            int firstRow = 0;      // zero‑based index
            int firstColumn = 0;
            int totalRows = 5;
            int totalColumns = 3;

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;

            // Add a simple comment to the first cell of the table (A1)
            int commentIdx = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Reference Document";   // set comment text
            comment.Author = "Automation";
            comment.IsVisible = true;

            // Save the workbook
            string outputPath = "TableWithComment.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
