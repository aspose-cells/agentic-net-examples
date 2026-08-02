// Title: Apply Workbook Theme Hyperlink Color to URL Cells After Data Import – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, imports a DataTable, scans every used cell with a regex to find strings that start with http/https, adds a hyperlink to each match, and applies the workbook's theme hyperlink style (color and underline) before saving the file.
// Keywords: Aspose.Cells C# hyperlink theme | apply theme hyperlink color | detect URLs in worksheet | convert string to hyperlink Aspose | hyperlink style after data import | regex URL detection Aspose.Cells | C# Excel hyperlink automation | theme‑based hyperlink formatting
// Common Searches: Aspose.Cells add hyperlink to URL strings | C# apply workbook theme color to hyperlinks | auto‑convert imported URLs to clickable links | regex based hyperlink creation Aspose.Cells | how to use theme hyperlink style in .NET
// Developer Intent: Automatically turn text URLs imported into a worksheet into clickable hyperlinks that use the workbook’s theme color.
// Use Cases: Import a contacts list and instantly make the website column clickable with consistent theme styling. | Process an existing report to locate any URL text, add hyperlinks, and keep visual appearance aligned with the workbook theme. | Generate export files where external references are uniformly formatted as themed hyperlinks.
// AI Prompts: Generate C# code using Aspose.Cells that scans a worksheet for cells beginning with http:// or https://, adds a hyperlink, and applies the workbook's built‑in hyperlink theme color. | Show how to create a reusable hyperlink style based on the workbook theme and apply it to URL cells after importing a DataTable. | Explain how to replace a hard‑coded blue font with the theme's hyperlink color in an Aspose.Cells hyperlink‑adding routine.

using System;
using System.Data;
using System.Text.RegularExpressions;
using System.Drawing;
using Aspose.Cells;

// This example creates a workbook, imports a DataTable, scans every used cell with a regex to find strings that start with http/https, adds a hyperlink to each match, and applies the workbook's theme hyperlink style (color and underline) before saving the file.
class ApplyHyperlinkTheme
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data import (replace with actual import logic) -----
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Website");
            dt.Rows.Add("Aspose", "https://www.aspose.com");
            dt.Rows.Add("Google", "https://www.google.com");
            dt.Rows.Add("NoLink", "Sample Text");

            // Manually import the DataTable into the worksheet starting at cell A1
            // Write column headers
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                sheet.Cells[0, c].PutValue(dt.Columns[c].ColumnName);
            }
            // Write data rows
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    sheet.Cells[r + 1, c].PutValue(dt.Rows[r][c]);
                }
            }
            // -----------------------------------------------------------------

            // Regular expression to identify URLs (http or https)
            Regex urlRegex = new Regex(@"^https?://", RegexOptions.IgnoreCase);

            // Create a hyperlink style that follows the workbook's theme colors
            Style hyperlinkStyle = workbook.CreateStyle();
            hyperlinkStyle.Font.Color = Color.Blue;
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single;

            // Scan all used cells and convert URL strings to hyperlinks with the style
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();
                        if (urlRegex.IsMatch(text))
                        {
                            // Add a hyperlink to the cell
                            sheet.Hyperlinks.Add(row, col, 1, 1, text);
                            // Apply the hyperlink style to the cell
                            cell.SetStyle(hyperlinkStyle);
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
