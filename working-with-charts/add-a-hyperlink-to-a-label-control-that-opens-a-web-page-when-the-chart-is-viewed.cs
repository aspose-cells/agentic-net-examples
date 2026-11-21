// Title: Add a hyperlink to a label shape in an Aspose.Cells chart using C#
// Common Searches: C# Aspose.Cells add hyperlink to worksheet label shape | How to make a label in an Excel chart clickable with Aspose.Cells | Set URL on label shape for Aspose.Cells generated chart in C# | Aspose.Cells chart label hyperlink example code | Create Excel chart with clickable label using Aspose.Cells C#
// Tags: Aspose.Cells label shape hyperlink C# | worksheet shape hyperlink Aspose.Cells | Excel chart clickable label Aspose.Cells | C# set hyperlink address on shape | Aspose.Cells chart label URL example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Demonstrates how to insert a label shape into a worksheet, assign a hyperlink URL to the label, and associate it with a chart using Aspose.Cells in C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a sample column chart (optional)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
                Chart chart = worksheet.Charts[chartIndex];

                // Add a label shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                Label label = worksheet.Shapes.AddLabel(0, 0, 0, 0, 200, 30);
                label.Text = "Click here to open website";

                // Set hyperlink for the label shape (URL is inferred from the address)
                label.Hyperlink.Address = "https://www.example.com";

                // Save the workbook
                string outputPath = "ChartWithHyperlink.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
