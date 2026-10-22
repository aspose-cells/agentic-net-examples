// Title: Use Aspose.Cells for .NET to find email addresses in Excel cells and replace them with mailto hyperlinks
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, scans each used cell for email patterns using a regular expression, and adds a mailto hyperlink to any cell containing an email address. | Update an existing Aspose.Cells workbook so that cells containing email addresses are overwritten with the plain email text and a clickable mailto link is inserted.
// Common Searches: aspnet c# Aspose.Cells add mailto hyperlink to cells containing email addresses | regex email detection in Excel workbook using Aspose.Cells .NET | how to replace cell text with email hyperlink in Aspose.Cells | iterate over used range and create mailto links in Aspose.Cells C# | convert email strings to clickable links in Excel file with Aspose.Cells
// Tags: email regex detection Aspose.Cells | hyperlink creation for email Aspose.Cells | iterate used cells Aspose.Cells C# | replace cell value with hyperlink Aspose.Cells | process .xlsx workbook Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an input.xlsx workbook, iterates through all used cells, extracts each cell's displayed string, uses a regular expression to locate email addresses, replaces the cell content with the email address, adds a mailto hyperlink to that cell, and saves the result as output.xlsx.
class EmailHyperlinkProcessor
{
    static void Main()
    {
        // Load existing workbook (use provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Regex pattern for email addresses
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", RegexOptions.Compiled);

        // Iterate through all used cells
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                Cell cell = cells[row, col];

                // Retrieve the formatted text of the cell
                string cellText = cell.StringValue; // formatted value as displayed

                // Find email addresses in the cell text
                Match match = emailRegex.Match(cellText);
                if (match.Success)
                {
                    string email = match.Value;

                    // Replace cell content with the email (if it wasn't already exactly the email)
                    cell.PutValue(email);

                    // Add a mailto hyperlink to the cell (use provided save rule later)
                    // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
                    sheet.Hyperlinks.Add(row, col, 1, 1, "mailto:" + email);
                }
            }
        }

        // Save the modified workbook (use provided save rule)
        workbook.Save("output.xlsx");
    }
}
