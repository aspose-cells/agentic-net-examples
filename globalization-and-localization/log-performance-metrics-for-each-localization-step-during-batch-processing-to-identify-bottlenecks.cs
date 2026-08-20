// Title: Measure and Log Localization Step Performance in Aspose.Cells Batch Processing (C#)
// Description: C# example that loads a template workbook, processes smart markers (localization), optionally calculates formulas, and saves the result while using Stopwatch to output elapsed milliseconds for loading, smart‑marker processing, formula calculation, and saving.
// Keywords: Aspose.Cells | C# | performance logging | smart markers timing | localization profiling | batch workbook processing | Stopwatch measurement | execution time tracking | formula calculation duration | save operation latency
// Common Searches: Aspose.Cells how to time smart marker processing | C# measure workbook load and save duration Aspose.Cells | profile localization step performance Aspose.Cells | batch processing performance metrics Aspose.Cells | stopwatch timing example Aspose.Cells
// Developer Intent: Capture elapsed time for each stage of a workbook localization batch to identify performance bottlenecks.
// Use Cases: Detect slow template loading in large‑scale localization jobs. | Pinpoint delays in smart‑marker processing to optimize localization speed. | Assess formula calculation overhead when generating localized reports. | Compare save performance across different file formats or storage locations. | Aggregate timing data from many workbooks to produce average and percentile metrics.
// AI Prompts: Generate C# code that writes each Stopwatch measurement to a CSV file for later analysis. | Show how to integrate the timing data with Serilog or Azure Application Insights in a batch workflow. | Provide a script that runs the example on multiple workbooks and outputs aggregated performance statistics. | Explain how to visualize the collected metrics in Power BI or Grafana dashboards. | Suggest ways to parallelize workbook processing while preserving per‑file timing information.

using System;
using System.Diagnostics;
using Aspose.Cells;

// C# example that loads a template workbook, processes smart markers (localization), optionally calculates formulas, and saves the result while using Stopwatch to output elapsed milliseconds for loading, smart‑marker processing, formula calculation, and saving.
class Program
{
    static void Main()
    {
        // -------------------- Load Workbook --------------------
        Stopwatch loadTimer = Stopwatch.StartNew();
        // Create (load) a workbook from a template file
        Workbook workbook = new Workbook("Template.xlsx");
        loadTimer.Stop();
        Console.WriteLine($"Load time: {loadTimer.ElapsedMilliseconds} ms");

        // -------------------- Process Smart Markers (Localization Step) --------------------
        Stopwatch processTimer = Stopwatch.StartNew();
        // Set up WorkbookDesigner for smart marker processing
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Sample data source for demonstration
        var sampleData = new[]
        {
            new { Product = "Widget", Quantity = 5, Price = 12.99 },
            new { Product = "Gadget", Quantity = 3, Price = 23.50 }
        };

        // Bind data source to the designer
        designer.SetDataSource("Items", sampleData);

        // Process the smart markers (this is the localization step we want to measure)
        designer.Process();
        processTimer.Stop();
        Console.WriteLine($"Smart marker processing time: {processTimer.ElapsedMilliseconds} ms");

        // -------------------- Calculate Formulas (Optional) --------------------
        Stopwatch calcTimer = Stopwatch.StartNew();
        workbook.CalculateFormula();
        calcTimer.Stop();
        Console.WriteLine($"Formula calculation time: {calcTimer.ElapsedMilliseconds} ms");

        // -------------------- Save Workbook --------------------
        Stopwatch saveTimer = Stopwatch.StartNew();
        workbook.Save("Result.xlsx");
        saveTimer.Stop();
        Console.WriteLine($"Save time: {saveTimer.ElapsedMilliseconds} ms");
    }
}
