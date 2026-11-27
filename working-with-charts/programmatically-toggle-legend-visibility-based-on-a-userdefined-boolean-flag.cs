// Title: How to toggle the legend visibility of an Excel chart with Aspose.Cells for .NET using a command‑line boolean
// Common Searches: aspnet hide chart legend based on command line argument using Aspose.Cells | c# set Legend.IsDeleted true or false in an Excel workbook with Aspose.Cells | programmatically show or hide Excel chart legend from a boolean flag | Aspose.Cells example to control chart legend visibility in a .xlsx file
// Tags: Aspose.Cells toggle chart legend visibility | Legend.IsDeleted API C# | Excel chart manipulation Aspose.Cells .xlsx | conditional chart legend display C# | programmatic chart formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing workbook or creates a new one, ensures a column chart exists, reads a boolean flag from the command line, sets the chart's Legend.IsDeleted property to hide or show the legend accordingly, and saves the modified workbook as a new .xlsx file.
class Program
{
    static void Main(string[] args)
    {
        // ------------------------------------------------------------
        // Determine the desired legend visibility from the command line.
        // Pass "true" to show the legend, "false" to hide it.
        // ------------------------------------------------------------
        bool showLegend = false; // default: hide
        if (args.Length > 0)
        {
            if (!bool.TryParse(args[0], out showLegend))
            {
                Console.WriteLine("Invalid argument. Use 'true' to show the legend or 'false' to hide it.");
                return;
            }
        }
        else
        {
            Console.WriteLine("No argument supplied. Defaulting to hide the legend.");
        }

        // ------------------------------------------------------------
        // File paths – adjust as needed.
        // ------------------------------------------------------------
        const string inputPath = "InputChart.xlsx";
        const string outputPath = "OutputChart.xlsx";

        Workbook workbook = null;

        try
        {
            // ------------------------------------------------------------
            // Load an existing workbook if it exists; otherwise create a sample.
            // ------------------------------------------------------------
            if (System.IO.File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
            }
            else
            {
                workbook = CreateSampleWorkbook();
                Console.WriteLine($"\"{inputPath}\" not found – created a sample workbook in memory.");
            }

            // ------------------------------------------------------------
            // Work with the first worksheet and its first chart.
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            if (sheet.Charts.Count == 0)
            {
                // Add a simple column chart if none exist.
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                Console.WriteLine("Added a default column chart.");
            }

            Chart targetChart = sheet.Charts[0];
            Legend legend = targetChart.Legend;

            // ------------------------------------------------------------
            // Toggle legend visibility.
            // Legend.IsDeleted = true  -> legend is hidden.
            // Legend.IsDeleted = false -> legend is visible.
            // ------------------------------------------------------------
            legend.IsDeleted = !showLegend;
            Console.WriteLine($"Legend visibility set to {(showLegend ? "visible" : "hidden")}.");

            // ------------------------------------------------------------
            // Save the modified workbook.
            // ------------------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }

    // ------------------------------------------------------------
    // Helper: creates a workbook with sample data and a chart.
    // ------------------------------------------------------------
    private static Workbook CreateSampleWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Header
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");

        // Sample rows
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("C");
        cells["B4"].PutValue(30);

        // Add a column chart linked to the data.
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        return wb;
    }
}
