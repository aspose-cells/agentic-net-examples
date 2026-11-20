// Title: Import a 2D string array into an Excel worksheet and format the header row with a bold blue font using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import two dimensional string array into worksheet | How to apply bold blue font to header row after importing data with Aspose.Cells | Set custom style for first row in Excel using Aspose.Cells .NET | C# code to import string[,] and style header cells in Excel | Aspose.Cells example for styling header row after data import
// Tags: import two-dimensional string array Aspose.Cells | apply custom header style Excel C# | bold blue font style Aspose.Cells | style first worksheet row Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The example creates a new Workbook, imports a two‑dimensional string array starting at cell A1, defines a style with a bold 12‑point blue font, applies this style to each cell in the first row (the header), and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Two‑dimensional array of strings (including header row)
        string[,] data = new string[,]
        {
            { "Header1", "Header2", "Header3" },
            { "Row1Col1", "Row1Col2", "Row1Col3" },
            { "Row2Col1", "Row2Col2", "Row2Col3" }
        };

        // Import the array into the worksheet starting at cell A1 (row 0, column 0)
        sheet.Cells.ImportTwoDimensionArray(data, 0, 0);

        // Create a custom style for the header row
        Style headerStyle = workbook.CreateStyle();
        Font headerFont = headerStyle.Font;
        headerFont.IsBold = true;          // Bold text
        headerFont.Size = 12;              // Font size 12
        headerFont.Color = Color.Blue;    // Blue color

        // Apply the style to each cell in the first row (header)
        int columnCount = data.GetLength(1);
        for (int col = 0; col < columnCount; col++)
        {
            sheet.Cells[0, col].SetStyle(headerStyle);
        }

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
