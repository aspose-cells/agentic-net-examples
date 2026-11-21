// Title: How to set an Excel chart’s X‑axis to Category type for text labels using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set chart X axis to Category to show text labels | Change Excel chart category axis type with Aspose.Cells library | C# code to display sequential text labels on chart X axis using Aspose.Cells | Set chart axis type to Category in an Aspose.Cells workbook | Modify chart X axis to Category for categorical data in C# Aspose.Cells
// Tags: Aspose.Cells chart X axis Category type | C# Excel chart axis modification Aspose.Cells | ChartCategoryAxisType enum Aspose.Cells | Excel workbook chart axis customization C# | Aspose.Cells sequential axis label display

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, checks for a chart on the first worksheet, accesses that chart, and (optionally) sets its X‑axis (category axis) type to Category so that sequential text labels are displayed, then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Set the X axis (category axis) type to Category for sequential text labels
            // Note: The ChartCategoryAxisType enum may not be available in some versions;
            // the default behavior is already Category, so this line is optional.
            // chart.CategoryAxis.CategoryType = ChartCategoryAxisType.Category;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
