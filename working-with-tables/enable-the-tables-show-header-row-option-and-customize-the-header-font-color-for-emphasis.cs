// Title: Show Table Header Row and Apply Red Bold Font with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, defines a ListObject that includes the header row, enables ShowHeaderRow, builds a custom TableStyle that sets the header font to red and bold, applies the style to the table, and saves the file as TableHeaderDemo.xlsx.
// Keywords: Aspose.Cells | C# | ShowHeaderRow | ListObject header | custom TableStyle | HeaderRow font color | red bold header | Excel table styling | .NET spreadsheet API
// Common Searches: Aspose.Cells show header row C# | How to style table header in Aspose.Cells | Custom TableStyle for header row Aspose.Cells .NET | Set header font color red Aspose.Cells | Enable ListObject header row programmatically
// Developer Intent: Display the table's header row and emphasize column titles with a red bold font using a custom TableStyle in Aspose.Cells for .NET.
// Use Cases: Activate the header row for a newly created ListObject and highlight it with a red bold font. | Create a reusable TableStyle that consistently formats header rows across multiple tables in a workbook. | Toggle the ShowHeaderRow property at runtime while preserving custom header styling for user‑controlled view options.
// AI Prompts: Generate C# code that creates a ListObject with ShowHeaderRow enabled and applies a custom TableStyle that sets the header font to blue and italic using Aspose.Cells. | Provide an example of updating an existing table's header style to a green background with white bold text in Aspose.Cells for .NET. | Show how to copy a custom header style from one worksheet to another table within the same workbook using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderDemo
{
    // Creates a workbook, adds sample data, defines a ListObject that includes the header row, enables ShowHeaderRow, builds a custom TableStyle that sets the header font to red and bold, applies the style to the table, and saves the file as TableHeaderDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.5);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.75);

            // Add a table (ListObject) that includes the header row
            int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Ensure the header row is visible
            table.ShowHeaderRow = true;

            // Create a style for the header row (red, bold font)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.Red;
            headerStyle.Font.IsBold = true;

            // Create a custom table style and set the header row element style
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
            string customStyleName = "CustomHeaderStyle";
            int styleIdx = tableStyles.AddTableStyle(customStyleName);
            TableStyle customStyle = tableStyles[styleIdx];

            // Add HeaderRow element and apply the header style
            int elementIdx = customStyle.TableStyleElements.Add(TableStyleElementType.HeaderRow);
            customStyle.TableStyleElements[elementIdx].SetElementStyle(headerStyle);

            // Apply the custom style to the table
            table.TableStyleName = customStyleName;

            // Save the workbook
            workbook.Save("TableHeaderDemo.xlsx");
        }
    }
}
