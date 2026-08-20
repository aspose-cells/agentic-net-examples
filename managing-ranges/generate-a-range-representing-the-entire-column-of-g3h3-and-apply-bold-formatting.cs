// Title: Bold Entire Columns G and H Using a G3:H3 Range in Aspose.Cells for .NET
// Description: Demonstrates how to create a range for cells G3:H3, retrieve its EntireColumn range, and apply a bold font style with a StyleFlag, then save the workbook as BoldEntireColumns.xlsx.
// Keywords: Aspose.Cells | C# | EntireColumn | StyleFlag | bold font | range G3:H3 | column formatting | Excel automation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells apply bold to entire column | How to use EntireColumn property in C# | Create range G3:H3 Aspose.Cells | StyleFlag bold font Aspose.Cells .NET | Bold columns G and H programmatically
// Developer Intent: Apply a bold font style to the full columns that contain the G3:H3 range.
// Use Cases: Emphasize header columns G and H in a generated report. | Highlight specific data sections by bolding entire columns. | Create a template where selected columns are automatically styled for readability.
// AI Prompts: Show C# code that creates a G3:H3 range, gets its EntireColumn, and applies a bold style using Aspose.Cells. | Explain how to use StyleFlag to change only the font bold attribute for columns G and H. | Provide an example of adding background color together with bold font to the same columns in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a range for cells G3:H3, retrieve its EntireColumn range, and apply a bold font style with a StyleFlag, then save the workbook as BoldEntireColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range that covers cells G3:H3 (row index 2, column indexes 6 and 7)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            AsposeRange range = cells.CreateRange(2, 6, 1, 2);

            // Get the entire columns that contain the range (columns G and H)
            AsposeRange entireColumns = range.EntireColumn;

            // Define a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Specify that only the bold attribute should be applied
            StyleFlag flag = new StyleFlag();
            flag.FontBold = true;

            // Apply the bold style to the entire columns
            entireColumns.ApplyStyle(boldStyle, flag);

            // Save the workbook
            workbook.Save("BoldEntireColumns.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
