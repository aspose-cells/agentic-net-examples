// Title: Show ListObject header row and apply a light‑blue solid background style with bold font using Aspose.Cells for .NET
// AI Prompts: Create a ListObject table in a worksheet, enable its header row, define a solid light‑blue style with bold text, apply the style to the header row range, and save the workbook as an .xlsx file. | Generate a reusable header style in Aspose.Cells for .NET, apply it to the first row of a ListObject, ensure the header is visible, and export the workbook.
// Common Searches: asp.net how to display ListObject header row with Aspose.Cells | c# apply custom background color to Excel table header using Aspose.Cells | asp.net set header row style for ListObject table in Excel | c# Aspose.Cells predefined style for table header row | asp.net save workbook with styled ListObject header
// Tags: Aspose.Cells ListObject header styling | C# solid background for Excel table header | Aspose.Cells predefined style application | C# enable ListObject header row | Aspose.Cells save workbook with styled table

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, adds a ListObject covering cells A1:B3, makes the header row visible, defines a light‑blue solid style with bold font, applies this style to the header row, and saves the file as ListObjectWithCustomHeader.xlsx.
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

            // Populate sample data (including header)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Define the range for the list object (including header row)
            int firstRow = 0;        // zero‑based index for row 1
            int firstColumn = 0;     // zero‑based index for column A
            int totalRows = 3;       // header + 2 data rows
            int totalColumns = 2;    // columns A and B

            // Add a ListObject (table) to the worksheet; the method returns the index of the added table
            int listIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject list = sheet.ListObjects[listIndex];

            // Ensure the header row is visible
            list.ShowHeaderRow = true;

            // Create a predefined style for the header background
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true;

            // Apply the style to the header row range of the list object
            AsposeRange headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            StyleFlag flag = new StyleFlag { All = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // Save the workbook to a file
            workbook.Save("ListObjectWithCustomHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
