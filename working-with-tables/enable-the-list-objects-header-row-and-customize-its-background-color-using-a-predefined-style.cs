// Title: Show ListObject Header Row and Apply a Custom Header Background Style in Aspose.Cells for .NET
// Description: Creates a workbook, adds a ListObject with a visible header row, defines a solid light‑green background and bold dark‑blue font style, builds a custom TableStyle that applies this style to the header row, assigns the style to the table, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells ListObject header row | C# custom table header style | apply background color to table header Aspose.Cells | create TableStyle programmatically | show header row Aspose.Cells | .NET Excel table styling
// Common Searches: Aspose.Cells how to display ListObject header row | custom header background color for Excel table using Aspose.Cells C# | create and apply TableStyle to ListObject in .NET | set header row visibility and style Aspose.Cells | sample code for styling Excel table header with Aspose
// Developer Intent: Make the ListObject header row visible and format it with a custom background and font style.
// Use Cases: Generate product catalogs where the header row uses corporate colors for instant visual distinction. | Define reusable table styles that enforce consistent header formatting across automated reports. | Produce Excel exports that comply with branding guidelines by highlighting table headers in .NET applications.
// AI Prompts: Show how to replace the custom TableStyle with a built‑in Aspose.Cells style while keeping the header visible. | Provide code to set different font colors for the header row and the total row in the same ListObject. | Explain how to apply the same custom header style to multiple ListObjects across several worksheets in one workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a ListObject with a visible header row, defines a solid light‑green background and bold dark‑blue font style, builds a custom TableStyle that applies this style to the header row, assigns the style to the table, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(2.5);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(1.2);

        // Add a ListObject (table) that includes the header row
        int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Ensure the header row is visible
        table.ShowHeaderRow = true;

        // Create a style for the header row (solid light‑green background, bold dark‑blue font)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.BackgroundColor = Color.LightGreen;
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.DarkBlue;

        // Create a custom table style and set the HeaderRow element to use the style above
        string customStyleName = "MyHeaderStyle";
        TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
        int styleIdx = tableStyles.AddTableStyle(customStyleName);
        TableStyle customTableStyle = tableStyles[styleIdx];

        // Add the HeaderRow element and assign the prepared style
        int elementIdx = customTableStyle.TableStyleElements.Add(TableStyleElementType.HeaderRow);
        TableStyleElement headerElement = customTableStyle.TableStyleElements[elementIdx];
        headerElement.SetElementStyle(headerStyle);

        // Apply the custom table style to the ListObject
        table.TableStyleName = customStyleName;

        // Save the workbook
        workbook.Save("ListObjectHeaderCustomStyle.xlsx");
    }
}
