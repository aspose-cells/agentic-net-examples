// Title: Show a table’s header row and color the header font red using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to display the header row of a ListObject | Change Excel table header font color to red with Aspose.Cells .NET | Apply a custom style to the header row of a table using Aspose.Cells in C# | Set ShowHeaderRow property and style header font in Aspose.Cells workbook | C# Aspose.Cells example for formatting table header row
// Tags: Aspose.Cells ListObject header visibility | Aspose.Cells style flag font color | C# Excel table header styling with Aspose.Cells | Aspose.Cells apply custom header style | Excel table header formatting Aspose.Cells .NET | Aspose.Cells create style for table header row

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The C# program creates a new workbook, inserts sample data, adds a ListObject (Excel table) with the header row shown, defines a style that sets the header font color to red, applies the style to the header row using a StyleFlag, and saves the file as TableWithHeaderStyle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with a header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Define the range for the table (including header row)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 2;      // two columns (Name, Age)

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Ensure the header row is shown
            table.ShowHeaderRow = true;

            // Create a style for the header font color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.Red; // Emphasis color

            // Apply the style only to the header row
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            StyleFlag flag = new StyleFlag
            {
                FontColor = true
            };
            headerRange.ApplyStyle(headerStyle, flag);

            // Save the workbook
            string outputPath = "TableWithHeaderStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
