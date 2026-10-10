// Title: Modify the first chart series data range and save the workbook as XLSX with Aspose.Cells for .NET
// AI Prompts: Load an existing Excel file using Aspose.Cells, set the first chart's series Values to the range A2:A5 and XValues to B2:B5, then save the workbook as an XLSX file to keep all chart formatting. | Using C# and Aspose.Cells, open a workbook, update the data source of the first chart series, and export the workbook to XLSX while preserving chart appearance and other formatting.
// Common Searches: c# aspnet change chart series range in existing Excel file using Aspose.Cells | how to preserve chart formatting when saving workbook as xlsx with Aspose.Cells | update first chart series values and categories in Aspose.Cells .NET example | Aspose.Cells save workbook after modifying chart data source to XLSX
// Tags: chart series range update Aspose.Cells | save workbook as XLSX preserving charts | modify chart data source C# Aspose.Cells | Aspose.Cells chart formatting retention | load existing Excel workbook edit chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, updates the first chart's first series Values and XValues ranges, and saves the workbook as an XLSX file to retain all formatting and chart details.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook (XLSX, XLS, etc.)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (sheet.Charts.Count > 0)
            {
                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Modify the first series if it exists
                if (chart.NSeries.Count > 0)
                {
                    // Set new data source for the series values (Y values)
                    chart.NSeries[0].Values = "A2:A5";

                    // Optionally set new X values (categories)
                    chart.NSeries[0].XValues = "B2:B5";
                }

                // Additional chart customizations can be added here
                // e.g., chart.Title.Text = "Updated Chart Title";
            }

            // Save the workbook as XLSX to preserve all formatting and chart details
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
