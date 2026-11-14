// Title: Create a merged title row with centered bold styling and a styled header in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge cells across columns for a table title | apply centered bold style to merged cells in Excel with Aspose.Cells | create a styled header row in an Aspose.Cells worksheet | save a workbook that contains a merged header row using Aspose.Cells C# | set font size and alignment for a merged title cell with Aspose.Cells
// Tags: merge cells Aspose.Cells C# | styled header row Aspose.Cells | centered bold title Excel Aspose | create workbook Aspose.Cells .NET | apply style to range Aspose.Cells

using Aspose.Cells;
using System;

// The example demonstrates how to generate an Excel workbook with Aspose.Cells for .NET, merge cells A1 through D1 to create a centered, bold, larger‑font title, add a styled header row (Product, Region, Quantity, Revenue), populate sample data rows, and save the file as TableWithMergedHeader.xlsx, with exception handling.
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

            // Set the title text in the first cell
            sheet.Cells["A1"].PutValue("Sales Report");

            // Merge cells A1 to D1 to span the title across four columns
            sheet.Cells.Merge(0, 0, 1, 4); // row, column, totalRows, totalColumns

            // Apply style to the merged title (centered, bold, larger font)
            Style titleStyle = workbook.CreateStyle();
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = 14;
            sheet.Cells["A1"].SetStyle(titleStyle);

            // Create header row in the second row
            sheet.Cells["A2"].PutValue("Product");
            sheet.Cells["B2"].PutValue("Region");
            sheet.Cells["C2"].PutValue("Quantity");
            sheet.Cells["D2"].PutValue("Revenue");

            // Apply style to header cells (centered, bold)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.Font.IsBold = true;

            // Apply the style to the range A2:D2
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A2:D2");
            StyleFlag flag = new StyleFlag() { All = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // Example data rows
            sheet.Cells["A3"].PutValue("Widget");
            sheet.Cells["B3"].PutValue("North");
            sheet.Cells["C3"].PutValue(120);
            sheet.Cells["D3"].PutValue(3000);

            sheet.Cells["A4"].PutValue("Gadget");
            sheet.Cells["B4"].PutValue("South");
            sheet.Cells["C4"].PutValue(85);
            sheet.Cells["D4"].PutValue(2125);

            // Save the workbook to a file
            string outputPath = "TableWithMergedHeader.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
