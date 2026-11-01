// Title: Show a table’s header row and apply a custom font color to the header using Aspose.Cells for .NET
// AI Prompts: Create a workbook, add a ListObject with headers, set ShowHeaderRow = true, and apply a dark‑red font color to the header row. | Given an existing worksheet, enable the table header row visibility and change its font color to a specified value using Aspose.Cells APIs. | Generate an Excel file where the table’s header row is displayed and styled with a custom font color without affecting the data rows.
// Common Searches: Aspose.Cells how to display table header row and change its font color in C# | C# Aspose.Cells set ShowHeaderRow true and apply font color to table header | apply custom font color to Excel table header using Aspose.Cells .NET | enable header row for ListObject and style header text in Aspose.Cells
// Tags: ListObject.ShowHeaderRow property | Aspose.Cells header row styling | apply font color to Excel table header C# | customize table header appearance Aspose.Cells | Excel table header visibility .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a workbook, adding a ListObject table with headers, enabling the header row, and applying a dark red font color to the header cells before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);

            // Add a table (ListObject) that includes the header and data rows
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Ensure the header row is visible
            table.ShowHeaderRow = true;

            // Determine the number of columns in the table
            int columnCount = table.EndColumn - table.StartColumn + 1;

            // Define the range that represents the header row of the table
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(table.StartRow, table.StartColumn, 1, columnCount);

            // Create a style with the desired font color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.DarkRed;

            // Apply only the font color change to the header range
            StyleFlag flag = new StyleFlag { FontColor = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // Save the workbook (save rule)
            string outputPath = "TableWithCustomHeader.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
