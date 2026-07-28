// Title: C# – Set a Custom RGB Fill Color for Header Cells with Aspose.Cells
// Description: Demonstrates how to create a workbook, define a header range (A1:D1), build a CellsColor using Color.FromArgb(0,128,128), configure a solid‑fill style, apply it with a StyleFlag that targets only cell shading, and save the file. Ideal for adding brand‑specific background colors to Excel headers in .NET applications.
// Keywords: Aspose.Cells C# fill color | custom RGB background Aspose.Cells | header row style Aspose.Cells | StyleFlag cell shading | solid pattern background .NET | Excel header color C#
// Common Searches: how to set RGB background color for a header row in Aspose.Cells C# | apply solid fill to a range of cells using Aspose.Cells .NET | use StyleFlag to change only cell shading in Aspose.Cells | create and reuse CellsColor with specific RGB values
// Developer Intent: Apply a specific RGB background color to header cells while preserving other formatting.
// Use Cases: Brand‑consistent Excel reports with teal or corporate‑color headers. | Highlight table headings in generated spreadsheets for better readability. | Separate sections of a worksheet by coloring distinct ranges with exact RGB values.
// AI Prompts: Write C# code that uses Aspose.Cells to set a solid RGB fill color for a given cell range, affecting only the background. | Show how to reuse a CellsColor object for multiple header rows with different RGB values in Aspose.Cells. | Explain the role of StyleFlag when applying a fill color without altering fonts or borders.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a header range (A1:D1), build a CellsColor using Color.FromArgb(0,128,128), configure a solid‑fill style, apply it with a StyleFlag that targets only cell shading, and save the file. Ideal for adding brand‑specific background colors to Excel headers in .NET applications.
    public class HeaderFillColorDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the header range (e.g., cells A1 to D1)
            Aspose.Cells.Range headerRange = worksheet.Cells.CreateRange("A1:D1");

            // Create a custom CellsColor using RGB values (e.g., teal color)
            CellsColor headerColor = workbook.CreateCellsColor();
            headerColor.Color = Color.FromArgb(0, 128, 128); // RGB(0,128,128)

            // Create a style and assign the custom fill color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = headerColor.Color; // Use the RGB color
            headerStyle.Pattern = BackgroundType.Solid;      // Solid fill pattern

            // Define a StyleFlag to apply only cell shading (fill color)
            StyleFlag flag = new StyleFlag
            {
                CellShading = true
            };

            // Apply the style to the header range
            headerRange.ApplyStyle(headerStyle, flag);

            // Optionally put some header text
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["C1"].PutValue("Header 3");
            worksheet.Cells["D1"].PutValue("Header 4");

            // Save the workbook
            string outputPath = "HeaderFillColorDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
