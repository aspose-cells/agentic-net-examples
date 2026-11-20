// Title: Set font size, color, and underline for an entire Excel column with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set font size for a whole column | how to change column font color using Aspose.Cells .NET | apply single underline to Excel column cells with Aspose.Cells | modify column style font attributes in Aspose.Cells C# example | save workbook after formatting column font in Aspose.Cells
// Tags: column font size Aspose.Cells C# | excel column font color Aspose.Cells | single underline column style Aspose.Cells | column style modification .NET | Aspose.Cells column formatting example

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, accesses column B, updates its style to set a 14‑point blue font with a single underline, applies the style to the entire column, and saves the file as ColumnFontSettings.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Specify the column index (e.g., column B = 1)
            int columnIndex = 1;

            // Get the column object
            Column column = sheet.Cells.Columns[columnIndex];

            // Retrieve the current style of the column
            Style columnStyle = column.GetStyle();

            // Configure font settings
            columnStyle.Font.Size = 14;                     // Font size
            columnStyle.Font.Color = Color.Blue;           // Font color
            columnStyle.Font.Underline = FontUnderlineType.Single; // Underline

            // Apply the modified style back to the column
            column.SetStyle(columnStyle);

            // Save the workbook to a file
            string outputPath = "ColumnFontSettings.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
