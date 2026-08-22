// Title: Detect email addresses in Excel cells and add mailto hyperlinks with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that scans every string cell in an Aspose.Cells workbook, uses a regular expression to locate email addresses, and inserts a 'mailto:' hyperlink for each detected address. | Demonstrate how to replace a cell's text with the first email match and attach a clickable mailto hyperlink using the Aspose.Cells API.
// Common Searches: how to create mailto hyperlinks from email addresses in an Excel file using Aspose.Cells C# | regex email detection in Aspose.Cells workbook and add hyperlink | iterate over cells in Aspose.Cells and convert email text to clickable link | Aspose.Cells replace cell text with first email and add hyperlink .NET
// Tags: email regex detection Aspose.Cells C# | add mailto hyperlink Aspose.Cells | process string cells Aspose.Cells workbook | replace cell value with email Aspose.Cells | hyperlink insertion Excel Aspose.Cells .NET

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsEmailHyperlinkDemo
{
    // The example creates a workbook, writes a string containing email addresses to a cell, iterates over all string cells, uses a regular expression to find email addresses, replaces the cell content with the first matched email, adds a clickable 'mailto:' hyperlink to that cell, and saves the file as EmailHyperlinks.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample cell containing one or more email addresses
            worksheet.Cells["A1"].PutValue("Contact: john.doe@example.com and jane@domain.org");

            // Regular expression to detect email addresses
            Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}");

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string cellText = cell.StringValue;
                    MatchCollection matches = emailRegex.Matches(cellText);

                    if (matches.Count > 0)
                    {
                        // For demonstration, replace the cell content with the first detected email
                        string firstEmail = matches[0].Value;
                        cell.PutValue(firstEmail);

                        // Add a clickable "mailto:" hyperlink to the same cell
                        worksheet.Hyperlinks.Add(cell.Name, 1, 1, "mailto:" + firstEmail);
                    }
                }
            }

            // Save the workbook with the added hyperlinks
            workbook.Save("EmailHyperlinks.xlsx");
        }
    }
}
