// Title: Show ListObject Header and Apply Custom Background Style with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a ListObject covering A1:B3, makes the header row visible, defines a solid LightBlue background style, builds a custom TableStyle for the HeaderRow, assigns the style to the table, and saves the file as ListObjectHeaderCustomStyle.xlsx.
// Keywords: Aspose.Cells ListObject header | custom table header background .NET | show header row Aspose.Cells | programmatic TableStyle Aspose.Cells | C# Excel table styling | solid fill table header Aspose.Cells
// Common Searches: Aspose.Cells change ListObject header background color | How to enable header row for a table in Aspose.Cells .NET | Create and assign a custom TableStyle in Aspose.Cells | Set solid fill for table header using Aspose.Cells C#
// Developer Intent: Display the ListObject’s header row and style its background with a custom color programmatically.
// Use Cases: Brand‑consistent Excel reports with colored table headers | Improve readability of generated spreadsheets by highlighting header rows | Reuse a predefined TableStyle across multiple worksheets in a workbook
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject, ensure the header row is visible, and apply a LightBlue solid background via a custom TableStyle. | Demonstrate how to create a TableStyle, set the HeaderRow element style, and assign it to a table in Aspose.Cells for .NET. | Show how to define a reusable custom table style and apply it to several worksheets with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a ListObject covering A1:B3, makes the header row visible, defines a solid LightBlue background style, builds a custom TableStyle for the HeaderRow, assigns the style to the table, and saves the file as ListObjectHeaderCustomStyle.xlsx.
    public class ListObjectHeaderCustomStyleDemo
    {
        public static void Run()
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
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(1.2);

                // Add a list object (table) covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Create a custom style for the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.BackgroundColor = Color.LightBlue;

                // Create a new table style and set the HeaderRow element style
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                int styleIndex = tableStyles.AddTableStyle("MyCustomStyle");
                TableStyle customTableStyle = tableStyles[styleIndex];
                TableStyleElementCollection elements = customTableStyle.TableStyleElements;

                // Add HeaderRow element and apply the custom style
                int elementIndex = elements.Add(TableStyleElementType.HeaderRow);
                TableStyleElement headerElement = elements[elementIndex];
                headerElement.SetElementStyle(headerStyle);

                // Assign the custom table style to the list object
                table.TableStyleName = "MyCustomStyle";

                // Save the workbook
                workbook.Save("ListObjectHeaderCustomStyle.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ListObjectHeaderCustomStyleDemo.Run();
        }
    }
}
