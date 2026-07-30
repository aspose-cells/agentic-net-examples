// Title: Apply Corporate Branding with a Custom Table Style in Aspose.Cells for .NET
// Description: Creates a workbook, fills sample data, defines a TableStyle named "CorporateBrandStyle" (header, first column, whole‑table formatting), adds a ListObject, applies the style, and saves as CorporateTableStyle.xlsx using the Aspose.Cells C# API.
// Keywords: Aspose.Cells | C# | custom TableStyle | Excel table branding | TableStyleElements | ListObject | header row formatting | first column style | table borders | CorporateBrandStyle
// Common Searches: Aspose.Cells create custom table style .NET | apply corporate colors to Excel table using Aspose.Cells | set header row style with TableStyleElements | format first column of a ListObject in C# | add thin gray borders to whole table Aspose.Cells | save styled workbook as xlsx with Aspose.Cells
// Developer Intent: Generate and apply a branded TableStyle to an Excel table programmatically.
// Use Cases: Standardize report tables to match company visual identity | Highlight key identifier column with bold text and brand‑color fill | Provide consistent thin gray grid lines across all table cells | Reuse the same style across multiple worksheets or workbooks
// AI Prompts: Write C# code that creates a TableStyle called 'CorporateBrandStyle' with a bold white header on a corporate blue background, a light‑blue first column, and thin gray borders for the whole table, then applies it to a ListObject. | Explain how to extend the custom TableStyle to include a total row with a distinct background and summary formulas. | Show how to export the workbook containing the styled table to PDF while preserving the custom TableStyle using Aspose.Cells. | Provide a step‑by‑step guide to modify the custom style colors via XML definition instead of code.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, fills sample data, defines a TableStyle named "CorporateBrandStyle" (header, first column, whole‑table formatting), adds a ListObject, applies the style, and saves as CorporateTableStyle.xlsx using the Aspose.Cells C# API.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (5 columns, 10 data rows)
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

            // Define a custom table style name that follows corporate branding
            string customStyleName = "CorporateBrandStyle";

            // Access the table style collection (rule: TableStyleCollection)
            TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;

            // Add a new custom table style (rule: AddTableStyle)
            int styleIndex = tableStyles.AddTableStyle(customStyleName);
            TableStyle customStyle = tableStyles[styleIndex];

            // ----- Header Row Style -----
            Style headerStyle = workbook.CreateStyle(); // rule: CreateStyle
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.FromArgb(0, 112, 192); // corporate blue

            // Set the style for the header row element (rule: SetElementStyle)
            TableStyleElement headerElement = customStyle.TableStyleElements[TableStyleElementType.HeaderRow];
            headerElement.SetElementStyle(headerStyle);

            // ----- First Column Style -----
            Style firstColStyle = workbook.CreateStyle();
            firstColStyle.Font.IsBold = true;
            firstColStyle.Pattern = BackgroundType.Solid;
            firstColStyle.ForegroundColor = Color.FromArgb(221, 235, 247); // light corporate blue

            TableStyleElement firstColElement = customStyle.TableStyleElements[TableStyleElementType.FirstColumn];
            firstColElement.SetElementStyle(firstColStyle);

            // ----- Whole Table Style (borders) -----
            Style wholeTableStyle = workbook.CreateStyle();
            wholeTableStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.BottomBorder].Color = Color.Gray;
            wholeTableStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.TopBorder].Color = Color.Gray;
            wholeTableStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.LeftBorder].Color = Color.Gray;
            wholeTableStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            wholeTableStyle.Borders[BorderType.RightBorder].Color = Color.Gray;

            TableStyleElement wholeTableElement = customStyle.TableStyleElements[TableStyleElementType.WholeTable];
            wholeTableElement.SetElementStyle(wholeTableStyle);

            // Create a table that covers the populated range
            int totalRows = 11; // header + 10 data rows
            int totalCols = 5;
            int tableIndex = sheet.ListObjects.Add(0, 0, totalRows - 1, totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply the custom style to the table (rule: TableStyleName property)
            table.TableStyleName = customStyleName;
            table.ShowTableStyleFirstColumn = true;
            // Header row is already displayed because the table was created with hasHeaders = true

            // Save the workbook (lifecycle rule)
            workbook.Save("CorporateTableStyle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
