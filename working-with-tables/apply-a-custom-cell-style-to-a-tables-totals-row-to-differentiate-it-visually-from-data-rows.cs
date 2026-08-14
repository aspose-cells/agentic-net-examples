// Title: C# Example: Apply a Custom Totals Row Style to an Aspose.Cells Table
// Description: Creates a workbook, adds a ListObject table with a summed totals row, defines a TableStyle that gives the totals row a light‑gray background, bold dark‑blue font, applies the style, and saves the file.
// Keywords: Aspose.Cells | C# | custom totals row style | TableStyle | ListObject | TableStyleElement TotalRow | background color totals row | bold font totals row | Excel table totals formatting | Aspose.Cells example
// Common Searches: Aspose.Cells change totals row style | C# set totals row background Aspose | How to format totals row in Aspose.Cells table | Create custom TableStyle for totals row .NET | Aspose.Cells sample totals row formatting
// Developer Intent: Style the totals row of a ListObject table so it stands out visually.
// Use Cases: Financial reports where the grand‑total row must be highlighted. | Inventory worksheets that need a distinct summary row. | Sales dashboards applying corporate brand colors to total rows. | Automated invoice generation with a clearly marked total line. | Multi‑sheet workbooks requiring consistent totals‑row formatting.
// AI Prompts: Write C# code using Aspose.Cells to create a TableStyle that formats the totals row with a light gray background and bold dark blue font. | Show how to apply the same custom totals‑row style to multiple tables in a workbook. | Explain how to modify the totals‑row style after the table has been created or after the workbook is saved. | Provide a step‑by‑step guide to add a totals row, set a sum calculation, and style it with Aspose.Cells. | Generate a GitHub‑ready snippet that demonstrates creating, styling, and saving a table with a custom totals row.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a ListObject table with a summed totals row, defines a TableStyle that gives the totals row a light‑gray background, bold dark‑blue font, applies the style, and saves the file.
class ApplyCustomTotalsRowStyle
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");

            // Populate data rows
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(15);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(12);

            // Add a table that includes the data range
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            // Set a display name for the table (Name property is not available in this version)
            table.DisplayName = "ProductsTable";

            // Show totals row and set a sum calculation for the Price column
            table.ShowTotals = true;
            ListColumn priceColumn = table.ListColumns[1]; // second column (Price)
            priceColumn.TotalsCalculation = TotalsCalculation.Sum;
            priceColumn.TotalsRowLabel = "Grand Total";

            // -----------------------------------------------------------------
            // Create a custom style for the totals row
            // -----------------------------------------------------------------
            Style totalsRowStyle = workbook.CreateStyle();
            totalsRowStyle.Pattern = BackgroundType.Solid;
            totalsRowStyle.ForegroundColor = Color.LightGray; // background color
            totalsRowStyle.Font.IsBold = true;                // bold font
            totalsRowStyle.Font.Color = Color.DarkBlue;       // font color

            // -----------------------------------------------------------------
            // Create a custom table style and assign the totals row style to it
            // -----------------------------------------------------------------
            string customStyleName = "CustomTotalsStyle";
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
            int styleIdx = tableStyles.AddTableStyle(customStyleName);
            TableStyle customTableStyle = tableStyles[styleIdx];

            // Add the TotalRow element and set its style
            int elementIdx = customTableStyle.TableStyleElements.Add(TableStyleElementType.TotalRow);
            TableStyleElement totalRowElement = customTableStyle.TableStyleElements[elementIdx];
            totalRowElement.SetElementStyle(totalsRowStyle);

            // Apply the custom table style to the table
            table.TableStyleName = customStyleName;

            // Save the workbook
            workbook.Save("TableWithCustomTotalsRowStyle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
