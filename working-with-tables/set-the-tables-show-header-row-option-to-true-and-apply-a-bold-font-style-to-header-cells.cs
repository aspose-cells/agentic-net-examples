// Title: Show Table Header Row and Apply Bold Font in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a ListObject with a visible header row, enables ShowHeaderRow, defines a bold font style, and applies it exclusively to the header row using StyleFlag before saving the file.
// Keywords: Aspose.Cells C# table header | ShowHeaderRow Aspose.Cells | Bold font style ListObject | StyleFlag font bold .NET | Excel table header formatting | Aspose.Cells Table styling | ListObject header visibility
// Common Searches: Aspose.Cells how to show table header row | C# set ShowHeaderRow true | Apply bold font to table header using Aspose.Cells | StyleFlag usage in Aspose.Cells | Format ListObject header in .NET
// Developer Intent: The developer needs to make a table’s header row visible and format it with bold text programmatically.
// Use Cases: Generate Excel reports where the table header must be clearly distinguished. | Automate workbook creation with consistent header styling across multiple sheets. | Integrate bold header formatting into data‑export pipelines that use Aspose.Cells.
// AI Prompts: Write C# code that adds a ListObject to a worksheet, sets ShowHeaderRow to true, and applies a bold font style to the header row using Aspose.Cells. | Show how to use StyleFlag to apply only the FontBold attribute to a specific row in Aspose.Cells. | Explain how to extend the example to add a background color and border to the table header while keeping ShowHeaderRow enabled.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderDemo
{
    // Creates a workbook, adds a ListObject with a visible header row, enables ShowHeaderRow, defines a bold font style, and applies it exclusively to the header row using StyleFlag before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header row and some sample data
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(1.8);

                // Add a table (ListObject) that includes the header row
                // Parameters: first row, first column, last row, last column, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Create a style with bold font for the header cells
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;

                // Apply the bold style only to the header row (row 0)
                // Use StyleFlag to limit the applied attributes to FontBold
                StyleFlag flag = new StyleFlag { FontBold = true };
                worksheet.Cells.ApplyRowStyle(0, boldStyle, flag);

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
