// Title: Verify that Aspose.Cells throws a CellsException when exporting a chart without a data source in C#
// Common Searches: C# Aspose.Cells test exporting a chart with no data source throws exception | How to catch CellsException from chart.ToImage when chart has no data source | Aspose.Cells chart export error handling without setting data source in .NET | Unit testing chart export failure in Aspose.Cells C# example | What exception is raised by Aspose.Cells when exporting an empty chart
// Tags: chart export exception Aspose.Cells | Aspose.Cells chart without data source | CellsException handling chart export | C# Aspose.Cells ToImage error | validate chart data source requirement

// Verify that exporting a chart without setting a data source throws an exception
using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart without assigning a data source, and attempts to export it with chart.ToImage. It demonstrates that Aspose.Cells raises a CellsException (or another exception) in this scenario, and shows how to catch and verify the error.
class ChartExportTest
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet (create rule)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a chart to the worksheet (create rule)
        // Chart type can be any; using Column as an example
        int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 15, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Do NOT set any data source for the chart

        try
        {
            // Attempt to export the chart to an image (export operation)
            // This should fail because the chart has no data source
            chart.ToImage("ChartWithoutData.png");
            // If no exception is thrown, the test has failed
            Console.WriteLine("Test Failed: No exception was thrown when exporting a chart without data source.");
        }
        catch (CellsException ex)
        {
            // Expected path: an exception is thrown
            Console.WriteLine("Test Passed: Caught expected CellsException.");
            Console.WriteLine("Exception Message: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Any other exception type is also acceptable for verification
            Console.WriteLine("Test Passed: Caught exception of type " + ex.GetType().Name);
            Console.WriteLine("Exception Message: " + ex.Message);
        }
    }
}
