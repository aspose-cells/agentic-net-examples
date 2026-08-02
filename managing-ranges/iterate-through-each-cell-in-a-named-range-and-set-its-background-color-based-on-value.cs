// Title: Color Cells in a Named Range with Aspose.Cells for .NET
// Description: Creates a workbook, fills A1:C5 with numbers, defines a named range "MyRange", iterates each cell in the range, assigns LightGreen, LightYellow, or LightCoral based on the cell's numeric value, applies a solid fill style, and saves the file as NamedRangeColoring.xlsx.
// Keywords: Aspose.Cells C# named range | iterate cells in named range | set cell background color Aspose.Cells | conditional fill based on value | Aspose.Cells style example | C# Excel heat map code | .NET Excel cell coloring
// Common Searches: Aspose.Cells loop through named range C# | how to change cell background color by value Aspose.Cells | apply conditional fill to a range using Aspose.Cells | retrieve Range object from named range Aspose.Cells .NET | C# code for heat‑map coloring in Excel with Aspose
// Developer Intent: The developer needs to walk through every cell in a predefined named range and programmatically set each cell’s background color according to its numeric value.
// Use Cases: Generate a heat‑map view for sales figures inside a specific data block. | Highlight low, medium, and high KPI values in a report without using Excel’s built‑in conditional formatting. | Apply consistent visual styling to a named area before exporting the workbook to end users.
// AI Prompts: Write C# code using Aspose.Cells to iterate a named range and color cells based on three numeric thresholds. | Show how to create a named range, loop through its cells, and apply a solid fill style with LightGreen, LightYellow, and LightCoral colors. | Explain how to reuse a Style object efficiently when coloring many cells in a named range with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, fills A1:C5 with numbers, defines a named range "MyRange", iterates each cell in the range, assigns LightGreen, LightYellow, or LightCoral based on the cell's numeric value, applies a solid fill style, and saves the file as NamedRangeColoring.xlsx.
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

            // Populate the worksheet with sample numeric data (A1:C5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue(row * 10 + col); // values: 0,1,2,...,42
                }
            }

            // Define a named range "MyRange" that refers to A1:C5
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!A1:C5";

            // Retrieve the Range object from the named range
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            Aspose.Cells.Range range = namedRange.GetRange(); // Resolve ambiguity with System.Range

            // Iterate through each cell in the range and set background color based on its value
            foreach (Cell cell in range)
            {
                double cellValue = cell.DoubleValue;
                Color bgColor;

                if (cellValue < 20)
                    bgColor = Color.LightGreen;
                else if (cellValue < 40)
                    bgColor = Color.LightYellow;
                else
                    bgColor = Color.LightCoral;

                // Create a style with solid fill and the chosen background color
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = bgColor;

                // Apply the style to the current cell
                cell.SetStyle(style);
            }

            // Save the workbook
            string outputPath = "NamedRangeColoring.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
