// Title: Convert the X‑axis of an existing Excel chart to a Category axis and assign a label with Aspose.Cells for .NET (C#)
// Common Searches: how to change Excel chart X axis to category type using Aspose.Cells C# | Aspose.Cells set category axis for an existing chart in .NET | C# code to modify chart axis type to category in an XLSX workbook | update X axis labels of a chart programmatically with Aspose.Cells | change chart axis to category and set its title using Aspose.Cells for .NET
// Tags: Aspose.Cells change chart X axis to category | C# set category axis title Aspose.Cells | modify existing Excel chart axis type .NET | programmatic chart axis conversion Aspose.Cells | Excel chart manipulation Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, accesses the first chart on the first worksheet, switches the X‑axis to a Category axis, sets its title to "Categories", and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = sheet.Charts[0];

            // Set the X axis (Category axis) title
            chart.CategoryAxis.Title.Text = "Categories";
            chart.CategoryAxis.Title.IsVisible = true;

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
