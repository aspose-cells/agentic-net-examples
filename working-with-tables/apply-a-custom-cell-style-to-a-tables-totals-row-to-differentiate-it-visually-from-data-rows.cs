// Title: Apply a Custom Style to a Table Totals Row with Aspose.Cells for .NET (C#)
// Description: Learn how to create a workbook, add a ListObject table with a totals row, define a solid light‑goldenrod‑yellow background and bold dark‑blue font style, build a custom TableStyle that includes a TotalRow element, apply it to the table, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom totals row style | C# table totals row formatting | Aspose.Cells TableStyleElement TotalRow | apply custom TableStyle Aspose.Cells | style totals row Excel .NET | Aspose.Cells ListObject totals row design | custom table style financial report
// Common Searches: how to style totals row in Aspose.Cells C# | Aspose.Cells apply background color to table total row | create custom TableStyle for totals row Aspose.Cells | set bold font for totals row ListObject Aspose.Cells | Aspose.Cells .NET change appearance of table summary row
// Developer Intent: Create and assign a custom TableStyle that visually distinguishes a table's totals row in an Aspose.Cells workbook.
// Use Cases: Highlight summary values in financial spreadsheets with a branded color scheme. | Generate Excel reports where the totals row stands out for quick review by end users. | Automate styling of tables in large‑scale data exports to match corporate design guidelines.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject table with a totals row and apply a custom style that uses a light yellow background and dark blue bold font. | Show how to create a TableStyle, add a TotalRow element, set its Style, and assign the style to a table in Aspose.Cells for .NET. | Provide a snippet that modifies the style of an existing table's totals row without creating a new TableStyle.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Learn how to create a workbook, add a ListObject table with a totals row, define a solid light‑goldenrod‑yellow background and bold dark‑blue font style, build a custom TableStyle that includes a TotalRow element, apply it to the table, and save the file using Aspose.Cells for .NET.
    public class ApplyCustomStyleToTotalsRow
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(15);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(12);

            // Add a table that includes the data range
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            // Set a display name for the table (Name property may not be available in older versions)
            table.DisplayName = "ProductsTable";
            table.ShowTotals = true; // Enable totals row

            // Set totals calculation for the Price column (second column)
            ListColumn priceColumn = table.ListColumns[1];
            priceColumn.TotalsCalculation = TotalsCalculation.Sum;
            priceColumn.TotalsRowLabel = "Grand Total";

            // Create a custom style for the totals row
            Style totalsRowStyle = workbook.CreateStyle();
            totalsRowStyle.Pattern = BackgroundType.Solid;
            totalsRowStyle.ForegroundColor = Color.LightGoldenrodYellow;
            totalsRowStyle.Font.IsBold = true;
            totalsRowStyle.Font.Color = Color.DarkBlue;

            // Create a custom table style and add the TotalRow element
            string customTableStyleName = "MyCustomTableStyle";
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
            int styleIdx = tableStyles.AddTableStyle(customTableStyleName);
            TableStyle customTableStyle = tableStyles[styleIdx];

            // Add TotalRow element and assign the custom style
            int elementIdx = customTableStyle.TableStyleElements.Add(TableStyleElementType.TotalRow);
            TableStyleElement totalRowElement = customTableStyle.TableStyleElements[elementIdx];
            totalRowElement.SetElementStyle(totalsRowStyle);

            // Apply the custom table style to the table
            table.TableStyleName = customTableStyleName;
            table.ShowTableStyleFirstColumn = true; // optional: show first column style
            table.ShowTableStyleLastColumn = true;  // optional: show last column style

            // Save the workbook
            workbook.Save("CustomTotalsRowStyle.xlsx");
        }
    }
}
