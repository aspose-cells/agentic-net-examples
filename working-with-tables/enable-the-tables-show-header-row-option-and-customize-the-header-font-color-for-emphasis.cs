using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using System.Drawing;

namespace AsposeCellsTableHeaderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate header row and some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.20);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.80);
            worksheet.Cells["A4"].PutValue("Orange");
            worksheet.Cells["B4"].PutValue(1.50);

            // Add a table (ListObject) that includes the header row
            // Parameters: first row, first column, last row, last column, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Ensure the header row is visible
            table.ShowHeaderRow = true;

            // Create a custom style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.Red;      // Emphasis color
            headerStyle.Font.IsBold = true;          // Optional bold for emphasis
            headerStyle.Font.Size = 12;              // Optional larger size

            // Define a custom table style name
            string customStyleName = "CustomHeaderStyle";

            // Access the collection of table styles in the workbook
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;

            // Add a new table style to the collection
            int styleIdx = tableStyles.AddTableStyle(customStyleName);
            TableStyle customTableStyle = tableStyles[styleIdx];

            // Access the elements collection of the new style
            TableStyleElementCollection elements = customTableStyle.TableStyleElements;

            // Add a HeaderRow element and assign the custom header style to it
            int headerElementIdx = elements.Add(TableStyleElementType.HeaderRow);
            elements[headerElementIdx].SetElementStyle(headerStyle);

            // Apply the custom style to the table
            table.TableStyleName = customStyleName;

            // Save the workbook to a file
            workbook.Save("TableWithHeaderStyle.xlsx");
        }
    }
}