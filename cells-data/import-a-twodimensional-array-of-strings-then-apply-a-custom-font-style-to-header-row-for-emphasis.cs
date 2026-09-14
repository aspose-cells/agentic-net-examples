// Title: Import a 2D string array into an Aspose.Cells workbook and style the header row with bold blue font in C#
// AI Prompts: Load a two‑dimensional object array into a worksheet starting at cell A1 using Cells.ImportTwoDimensionArray. | Create a Style with a 12‑point blue font set to bold, and configure a flag object to limit formatting to font properties. | Apply the style to the first row with Cells.ApplyRowStyle and write the workbook to an .xlsx file.
// Common Searches: how to import a 2d string array into Aspose.Cells worksheet C# | apply custom colored bold header row in Aspose.Cells | Aspose.Cells C# example for styling first row after importing data | import array and format header row in Aspose.Cells .NET | set font size and color for header row in Aspose.Cells workbook
// Tags: ImportTwoDimensionArray Aspose.Cells C# | ApplyRowStyle header row | StyleFlag font attributes Aspose.Cells | custom header font style workbook | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example imports a two‑dimensional string array into a new workbook, creates a bold blue 12‑point font style for the header row using a style flag, applies the style to the first row, and saves the file as TwoDimArrayWithHeaderStyle.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Two‑dimensional array of strings (including header row)
        object[,] data = new object[,]
        {
            { "Name", "Age", "City" },          // Header row
            { "Alice", "30", "New York" },
            { "Bob", "25", "Los Angeles" }
        };

        // Import the array starting at cell A1 (row 0, column 0)
        cells.ImportTwoDimensionArray(data, 0, 0);

        // Create a custom style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;               // Bold font
        headerStyle.Font.Color = Color.Blue;          // Blue color
        headerStyle.Font.Size = 12;                   // Font size

        // Define which style attributes to apply (only font related)
        StyleFlag flag = new StyleFlag
        {
            FontBold = true,
            FontColor = true,
            FontSize = true
        };

        // Apply the style to the first row (header)
        cells.ApplyRowStyle(0, headerStyle, flag);

        // Save the workbook
        workbook.Save("TwoDimArrayWithHeaderStyle.xlsx");
    }
}
