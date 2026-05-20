using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsCustomStyleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: simulate cells that are mapped to /Customer/Name
            // In a real scenario these cells would be populated via XML import/mapping
            string[] customerNames = { "Alice Johnson", "Bob Smith", "Carol Davis" };
            for (int i = 0; i < customerNames.Length; i++)
            {
                // Populate column A with customer names (starting from row 1)
                cells[i, 0].PutValue(customerNames[i]);
            }

            // Create a custom style to highlight customer names
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.ForegroundColor = Color.Yellow;      // Cell background
            highlightStyle.Font.IsBold = true;                  // Bold font
            highlightStyle.Font.Color = Color.DarkBlue;         // Font color
            highlightStyle.Font.Size = 12;                      // Font size

            // Apply the custom style to all cells in column A that contain a name
            for (int row = 0; row < customerNames.Length; row++)
            {
                Cell nameCell = cells[row, 0];
                // Ensure the cell actually has a value before styling
                if (!string.IsNullOrEmpty(nameCell.StringValue))
                {
                    nameCell.SetStyle(highlightStyle);
                }
            }

            // Save the workbook with the applied style
            workbook.Save("CustomerNamesHighlighted.xlsx");
        }
    }
}