// Title: Create and Apply a Custom XML‑Defined TableStyle in C# with Aspose.Cells for .NET
// Description: This example demonstrates how to generate a workbook, fill a 5‑column × 10‑row range, define a new TableStyle called "CorporateBrandStyle" using Aspose.Cells' TableStyleCollection, configure header, first‑column, and whole‑table elements (background colors, font attributes, thin borders), create a ListObject from the data range, apply the custom style, enable first‑column formatting, and save the result as CorporateTableStyle.xlsx.
// Keywords: Aspose.Cells | C# | .NET | custom TableStyle | XML defined table style | Excel table branding | ListObject styling | TableStyleCollection | programmatic Excel formatting | corporate Excel template
// Common Searches: Aspose.Cells create custom TableStyle C# | apply XML defined table style Aspose.Cells | set header row color in Aspose.Cells table | add thin borders to whole table Aspose.Cells | brand Excel tables with custom style using Aspose
// Developer Intent: Generate a reusable, brand‑compliant TableStyle in code and apply it to a worksheet table.
// Use Cases: Produce quarterly financial reports with a dark‑blue header and white bold text that matches corporate guidelines. | Export sales dashboards where the first column is highlighted in light gray for quick row identification. | Create standardized Excel templates that enforce thin borders and a uniform font size across all tables.
// AI Prompts: Write C# code that loads an XML‑defined TableStyle and assigns it to a ListObject using Aspose.Cells. | Show how to modify the HeaderRow element of an existing TableStyle to change its font size and background color. | Explain the steps to programmatically add a custom TableStyle, enable ShowTableStyleFirstColumn, and save the workbook with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example demonstrates how to generate a workbook, fill a 5‑column × 10‑row range, define a new TableStyle called "CorporateBrandStyle" using Aspose.Cells' TableStyleCollection, configure header, first‑column, and whole‑table elements (background colors, font attributes, thin borders), create a ListObject from the data range, apply the custom style, enable first‑column formatting, and save the result as CorporateTableStyle.xlsx.
class CorporateTableStyleDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (5 columns, 10 rows of data)
            for (int col = 0; col < 5; col++)
            {
                cells[0, col].PutValue($"Header {col + 1}");
            }
            for (int row = 1; row <= 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue(row * (col + 1));
                }
            }

            // ------------------------------------------------------------
            // Create a custom table style that follows corporate branding
            // ------------------------------------------------------------
            string styleName = "CorporateBrandStyle";

            // Access the table style collection and add a new style
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
            int styleIndex = tableStyles.AddTableStyle(styleName);
            TableStyle customStyle = tableStyles[styleIndex];
            TableStyleElementCollection elements = customStyle.TableStyleElements;

            // Header row style: dark blue background, white bold font
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.BackgroundColor = Color.DarkBlue;
            headerStyle.Font.Color = Color.White;
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 12;
            elements.Add(TableStyleElementType.HeaderRow);
            elements[TableStyleElementType.HeaderRow].SetElementStyle(headerStyle);

            // First column style: light gray background
            Style firstColStyle = workbook.CreateStyle();
            firstColStyle.Pattern = BackgroundType.Solid;
            firstColStyle.BackgroundColor = Color.LightGray;
            elements.Add(TableStyleElementType.FirstColumn);
            elements[TableStyleElementType.FirstColumn].SetElementStyle(firstColStyle);

            // Whole table style: thin borders and standard font size
            Style wholeTableStyle = workbook.CreateStyle();
            wholeTableStyle.Font.Size = 11;
            wholeTableStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            elements.Add(TableStyleElementType.WholeTable);
            elements[TableStyleElementType.WholeTable].SetElementStyle(wholeTableStyle);

            // ------------------------------------------------------------
            // Create a table from the data range and apply the custom style
            // ------------------------------------------------------------
            int tableIndex = sheet.ListObjects.Add(0, 0, 10, 4, true);
            ListObject table = sheet.ListObjects[tableIndex];

            table.TableStyleName = styleName;               // Apply custom style
            table.ShowTableStyleFirstColumn = true;         // Show first column formatting
            // Header row formatting is automatically applied; no explicit property needed.

            // Save the workbook with the styled table
            string outputPath = "CorporateTableStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
