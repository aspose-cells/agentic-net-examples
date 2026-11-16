// Title: Set a custom font family for every chart title in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change font family of all chart titles in a workbook | apply branding font to Excel chart titles using Aspose.Cells .NET | iterate worksheets and set chart title font name with Aspose.Cells | C# code to ensure each chart title uses a specific font in Aspose.Cells | bulk update chart title font family in existing Excel file Aspose.Cells
// Tags: set chart title font family Aspose.Cells | bulk chart title formatting C# Aspose.Cells | apply custom font to Excel chart titles .NET | iterate worksheets charts Aspose.Cells | ensure chart title exists Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, defines a custom font family, loops through every worksheet and each chart, creates a default title when missing, assigns the custom font to each chart title, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the custom font family to be applied
            string customFontFamily = "MyBrandFont";

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has a title; if not, set a default title
                    if (string.IsNullOrEmpty(chart.Title?.Text))
                    {
                        chart.Title.Text = "Chart Title";
                    }

                    // Apply the custom font family to the chart title
                    chart.Title.Font.Name = customFontFamily;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
