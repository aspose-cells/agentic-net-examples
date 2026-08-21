// Title: Aspose.Cells .NET: Show Table Header Row and Apply Bold Font to Header Cells
// Description: Demonstrates how to create a workbook, add a ListObject (Excel table) with a visible header row, and format that header row with a bold font using Style and StyleFlag in C#.
// Keywords: Aspose.Cells C# table header | ShowHeaderRow Aspose.Cells | bold header style Aspose.Cells | ListObject formatting .NET | Excel table styling Aspose
// Common Searches: Aspose.Cells show table header row | apply bold font to Excel table header C# | ListObject header formatting Aspose | StyleFlag bold header Aspose.Cells example | C# code to style table header in Aspose.Cells
// Developer Intent: Make the table header visible and emphasize it with bold text.
// Use Cases: Generate product catalogs where the column titles stand out for readers. | Export data to Excel with a pre‑styled table that meets corporate branding guidelines. | Create a reusable utility that adds a formatted table to any workbook for reporting pipelines.
// AI Prompts: Write C# code with Aspose.Cells that adds a ListObject, sets ShowHeaderRow to true, and applies a bold font only to the header row. | Show how to create a Style, enable Font.IsBold, and use StyleFlag to style the first row of a worksheet containing a table. | Provide an Aspose.Cells example that ensures the table header is displayed and formatted in bold before saving the file.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderDemo
{
    // Demonstrates how to create a workbook, add a ListObject (Excel table) with a visible header row, and format that header row with a bold font using Style and StyleFlag in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with a header row
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(1.20);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(0.80);

                // Add a table (ListObject) that includes the header row
                // Parameters: first row, first column, last row, last column, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Create a style with bold font for the header cells
                Style boldHeaderStyle = workbook.CreateStyle();
                boldHeaderStyle.Font.IsBold = true;

                // Apply the bold style only to the header row (row 0)
                // Use StyleFlag to limit the applied attributes to FontBold
                StyleFlag flag = new StyleFlag { FontBold = true };
                worksheet.Cells.ApplyRowStyle(0, boldHeaderStyle, flag);

                // Save the workbook
                workbook.Save("TableWithBoldHeader.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
