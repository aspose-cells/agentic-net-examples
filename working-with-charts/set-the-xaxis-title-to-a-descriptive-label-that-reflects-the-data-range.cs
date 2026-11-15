// Title: Programmatically set a visible X‑axis (category axis) title on an existing Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells set X axis title programmatically in C# | make category axis title visible on an Excel chart with Aspose.Cells | change existing chart X‑axis label using Aspose.Cells .NET | C# update chart axis title text without recreating the chart | Aspose.Cells chart axis formatting set title text
// Tags: set category axis title Aspose.Cells | make chart axis title visible .NET | update Excel chart X‑axis label C# | Aspose.Cells chart axis text property | modify existing chart axis Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing workbook, checks for a chart on the first worksheet, makes the category (X) axis title visible, assigns a descriptive label (e.g., "Date (Jan – Mar 2023)"), and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Get the first chart on the worksheet
                Chart chart = sheet.Charts[0];

                // Make the X‑axis (Category Axis) title visible
                chart.CategoryAxis.Title.IsVisible = true;

                // Set a descriptive label that reflects the data range
                chart.CategoryAxis.Title.Text = "Date (Jan – Mar 2023)";

                // Note: The Title class for axis does not expose an Overlay property in this API version.
                // If needed, additional formatting can be applied via the Font property.
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
