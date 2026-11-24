// Title: Show ListObject header row and apply a LightBlue background style in Aspose.Cells for .NET (C#)
// Common Searches: how to display the header row of a ListObject table using Aspose.Cells C# | apply a solid background color to an Excel table header with Aspose.Cells .NET | create and assign a custom style to a ListObject header row in C# | set cell shading for a ListObject header range when HeaderRowRange is unavailable | manual header range creation for ListObject tables in Aspose.Cells
// Tags: listobject header row enable aspose.cells | excel table header shading aspose.cells | define custom style foreground color aspose.cells | construct header range manually aspose.cells | c# aspose.cells listobject formatting

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, adds a ListObject with a visible header row, defines a LightBlue solid background style, builds the header range manually, applies the shading style to that range, and saves the file as ListObjectHeaderStyled.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including header row)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Define the range for the list object (table)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // including header
            int totalColumns = 3;

            // Add a ListObject (table) to the worksheet
            int listIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
            ListObject list = sheet.ListObjects[listIndex];
            list.ShowHeaderRow = true; // ensure header row is visible

            // Create a style for the header row background
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            // Define which style attributes to apply (only cell shading)
            StyleFlag styleFlag = new StyleFlag
            {
                All = false,
                CellShading = true
            };

            // Apply the style to the header row range of the table
            // Since ListObject.HeaderRowRange is unavailable in this version,
            // construct the header range manually.
            AsposeRange headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            headerRange.ApplyStyle(headerStyle, styleFlag);

            // Save the workbook
            string outputPath = "ListObjectHeaderStyled.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
