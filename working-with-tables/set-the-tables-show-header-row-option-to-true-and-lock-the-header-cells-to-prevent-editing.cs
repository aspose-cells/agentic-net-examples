using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header values and some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(2.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(1.8);

            // Create a table (ListObject) that includes the header row
            // Parameters: first row, first column, last row, last column, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Ensure the header row is visible
            table.ShowHeaderRow = true;

            // Lock the header cells so they cannot be edited
            for (int col = table.StartColumn; col <= table.EndColumn; col++)
            {
                Cell headerCell = sheet.Cells[table.StartRow, col];
                Style style = headerCell.GetStyle();
                style.IsLocked = true;               // Mark cell as locked
                headerCell.SetStyle(style);          // Apply the style back to the cell
            }

            // Protect the worksheet to enforce the locked cells
            sheet.Protect(ProtectionType.All);

            // Define output file path
            string outputPath = "TableHeaderLocked.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}