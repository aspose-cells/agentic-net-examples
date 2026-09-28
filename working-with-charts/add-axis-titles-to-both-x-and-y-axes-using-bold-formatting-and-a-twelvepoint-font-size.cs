// Title: Add bold 12‑point X‑axis and Y‑axis titles to an existing Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set the CategoryAxis.Title text, make it bold, and set the font size to 12 points for a chart in a workbook. | Update the first chart on the first worksheet to add formatted X and Y axis titles (bold, 12‑pt) and save the workbook with Aspose.Cells for .NET.
// Common Searches: how to set bold 12 point font for chart axis titles in Aspose.Cells C# | Aspose.Cells C# add X axis title to existing chart | change value axis title font size Aspose.Cells .NET | programmatically format Excel chart axis titles using Aspose.Cells | C# Aspose.Cells set category axis title text and style
// Tags: Aspose.Cells chart axis title formatting | set category axis title font Aspose.Cells | apply bold font to Excel chart axis Aspose.Cells | modify chart axis titles .NET | C# Aspose.Cells update existing chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, accesses the first chart, assigns custom text to the X (category) and Y (value) axes, applies bold 12‑point formatting to both titles, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // ----- X Axis (Category Axis) Title -----
            chart.CategoryAxis.Title.Text = "X Axis Title";
            chart.CategoryAxis.Title.Font.IsBold = true;
            chart.CategoryAxis.Title.Font.Size = 12;

            // ----- Y Axis (Value Axis) Title -----
            chart.ValueAxis.Title.Text = "Y Axis Title";
            chart.ValueAxis.Title.Font.IsBold = true;
            chart.ValueAxis.Title.Font.Size = 12;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
