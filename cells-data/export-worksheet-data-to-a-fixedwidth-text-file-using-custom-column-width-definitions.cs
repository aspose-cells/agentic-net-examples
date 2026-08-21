// Title: Export a Worksheet to a Fixed‑Width Text File with Custom Column Widths using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, set column widths in character units, and generate a UTF‑8 fixed‑width text file. Each cell value is truncated or right‑padded to match the defined width, producing a flat file suitable for legacy systems.
// Keywords: Aspose.Cells | C# | .NET | fixed width export | custom column width | Excel to text file | UTF‑8 flat file | column padding | cell truncation | legacy mainframe import
// Common Searches: Aspose.Cells export fixed width text C# | how to create fixed‑width file from Excel using Aspose | C# write fixed‑width flat file with column widths | truncate and pad Excel cells for fixed‑field export | Aspose.Cells set column width for text export
// Developer Intent: Produce a fixed‑width text file from an Excel worksheet where each column follows predefined character widths, handling both truncation and right‑padding automatically.
// Use Cases: Generate legacy mainframe input files that require exact column positions. | Create fixed‑field flat files for batch loading into older databases or ERP systems. | Automate production of aligned text reports for printing or archival without delimiters.
// AI Prompts: Write C# code with Aspose.Cells that exports a worksheet to a UTF‑8 fixed‑width text file using specified column widths, including truncation and right‑padding logic. | Provide a reusable method that accepts a Worksheet and an array of integer column widths and returns a fixed‑width string representation of the data. | Explain how Aspose.Cells column width units map to character counts and how to convert them for accurate fixed‑width file generation.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace FixedWidthExportExample
{
    // Demonstrates how to create a workbook, set column widths in character units, and generate a UTF‑8 fixed‑width text file. Each cell value is truncated or right‑padded to match the defined width, producing a flat file suitable for legacy systems.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Country");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue("United States");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue("Canada");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue("United Kingdom");

            // Define custom column widths (in character units)
            // Widths are chosen to accommodate the longest expected content
            sheet.Cells.SetColumnWidth(0, 5);   // ID column
            sheet.Cells.SetColumnWidth(1, 12);  // Name column
            sheet.Cells.SetColumnWidth(2, 15);  // Country column

            // Save the workbook in Excel format (lifecycle save)
            workbook.Save("Sample.xlsx");

            // Determine the range to export
            int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
            int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last column with data

            // Prepare a StringBuilder to collect all lines
            StringBuilder sb = new StringBuilder();

            // Iterate through each row
            for (int row = 0; row <= maxRow; row++)
            {
                StringBuilder line = new StringBuilder();

                // Iterate through each column
                for (int col = 0; col <= maxCol; col++)
                {
                    // Get the cell value as string; if the cell is null, treat as empty string
                    Cell cell = sheet.Cells[row, col];
                    string cellText = cell?.StringValue ?? string.Empty;

                    // Get the defined column width (in characters)
                    double colWidth = sheet.Cells.GetColumnWidth(col);

                    // Convert width to an integer number of characters for padding
                    int width = (int)Math.Floor(colWidth);

                    // Truncate if the text exceeds the column width
                    if (cellText.Length > width)
                    {
                        cellText = cellText.Substring(0, width);
                    }

                    // Pad the text on the right with spaces to achieve fixed width
                    line.Append(cellText.PadRight(width));
                }

                // Append the constructed line to the overall output
                sb.AppendLine(line.ToString());
            }

            // Write the fixed‑width text to a file
            File.WriteAllText("FixedWidthExport.txt", sb.ToString(), Encoding.UTF8);
        }
    }
}
