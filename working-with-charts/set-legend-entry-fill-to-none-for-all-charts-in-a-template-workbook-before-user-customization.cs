// Title: Set all chart legend entries to transparent text (no fill) in an Excel template using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set legend entry text no fill for every chart in a workbook | How to remove background color from chart legend entries using Aspose.Cells .NET | Iterate worksheets and charts to make legend text transparent with Aspose.Cells | C# code to disable legend entry fill in Excel charts via Aspose.Cells | Batch update chart legends to no fill in an Excel template with Aspose.Cells
// Tags: legendentry.istextnofill aspnet | aspocells modify chart legends | iterate worksheets charts aspocells | excel chart legend transparency c# | batch update chart legends aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SetLegendEntryFillNone
{
    // The example loads a template workbook, loops through each worksheet and its charts, sets LegendEntry.IsTextNoFill to true for every series, recalculates the charts, and saves the workbook, resulting in legend text with no fill across all charts.
    class Program
    {
        /// <param name="args">Command‑line arguments (optional).</param>
        static void Main(string[] args)
        {
            // Define input and output file names.
            // Adjust the paths as necessary for your environment.
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Load the template workbook.
                Workbook workbook = new Workbook(templatePath);
                Console.WriteLine($"Loaded workbook \"{templatePath}\" successfully.");

                // Iterate through each worksheet in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each chart on the worksheet.
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Ensure the chart has at least one series.
                        if (chart.NSeries == null || chart.NSeries.Count == 0)
                            continue;

                        // Iterate through each series (NSeries) of the chart.
                        for (int i = 0; i < chart.NSeries.Count; i++)
                        {
                            // Access the legend entry associated with the current series.
                            LegendEntry legendEntry = chart.NSeries[i].LegendEntry;

                            // Set the legend entry to have no fill for its text.
                            // According to Aspose.Cells documentation, IsTextNoFill = true
                            // means the text will have no fill (transparent).
                            legendEntry.IsTextNoFill = true;

                            // Optional: ensure the legend entry is not deleted.
                            legendEntry.IsDeleted = false;
                        }

                        // After modifying legend entries, you may want to recalculate the chart
                        // to ensure any dependent properties are refreshed.
                        chart.Calculate();
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as \"{outputPath}\" with legend entries set to no fill.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
