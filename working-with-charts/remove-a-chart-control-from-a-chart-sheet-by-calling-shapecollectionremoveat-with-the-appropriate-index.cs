// Title: How to delete a chart control from a chart sheet using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a workbook, adds a column chart to a chart sheet, then removes the chart with ShapeCollection.RemoveAt (or Charts.RemoveAt) and saves the file. | Generate a method that checks if a worksheet contains any charts and removes the chart at a given index using Aspose.Cells APIs. | Provide an example that prints the chart count before and after deleting the first chart control on a chart sheet with Aspose.Cells.
// Common Searches: aspocells c# remove chart from chart sheet example | how to delete a chart control in an Excel workbook using Aspose.Cells | C# Aspose.Cells ShapeCollection.RemoveAt chart index | remove first chart from worksheet programmatically Aspose.Cells | Aspose.Cells chart count before after removal C#
// Tags: Aspose.Cells chart removal C# | ShapeCollection.RemoveAt chart control | Excel chart sheet delete Aspose | C# chart collection manipulation | Aspose.Cells workbook chart count

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a worksheet named "MyChartSheet", inserts a column chart, displays the chart count, removes the chart at index 0 if present, shows the updated count, and saves the workbook as "RemovedChartControl.xlsx".
class RemoveChartControlFromChartSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet and obtain its reference
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            sheet.Name = "MyChartSheet";

            // Add a chart (chart control) to the worksheet
            // The Add method returns the index of the newly created chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex]; // optional reference to the chart object

            // (Optional) Set up chart data here if needed

            // Show the number of charts before removal
            Console.WriteLine("Charts count before removal: " + sheet.Charts.Count);

            // Remove the chart control at index 0, if any exist
            if (sheet.Charts.Count > 0)
            {
                sheet.Charts.RemoveAt(0);
            }

            // Show the number of charts after removal
            Console.WriteLine("Charts count after removal: " + sheet.Charts.Count);

            // Save the workbook to a file
            string outputPath = "RemovedChartControl.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
