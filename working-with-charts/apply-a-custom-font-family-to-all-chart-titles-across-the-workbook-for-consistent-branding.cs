// Title: How to apply a custom branding font family to every chart title in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, iterates all worksheets and charts, creates a visible title if missing, and sets the title's Font.Name to a specified brand font. | Write a C# method that takes a workbook path and a font family string, then updates all chart titles in the workbook to use that font via Aspose.Cells. | Provide a C# snippet that demonstrates changing the font family, size, and style of chart titles across multiple sheets with Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart title font family for all charts in a workbook | How to change Excel chart title font using Aspose.Cells .NET | Iterate through worksheets and charts to apply branding font with Aspose.Cells | Ensure chart titles are visible and assign a custom font in Aspose.Cells C# example
// Tags: Aspose.Cells chart title font customization | C# apply branding font to Excel chart titles | iterate worksheets charts Aspose.Cells | set chart title visibility Aspose.Cells | custom font family Excel chart titles .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartTitleFontExample
{
    // The program loads an existing Excel workbook, loops through each worksheet and its charts, ensures every chart has a visible title, applies a specified custom font family to each chart title, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Custom font family to apply to chart titles
                string customFontFamily = "MyBrandFont";

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        try
                        {
                            // Ensure the chart has a visible title; create one if missing
                            if (!chart.Title.IsVisible)
                            {
                                chart.Title.IsVisible = true;
                                chart.Title.Text = "Chart Title";
                            }

                            // Apply the custom font family to the chart title
                            chart.Title.Font.Name = customFontFamily;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
