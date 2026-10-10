// Title: Add data labels (series name, category name, value) to the first chart in an Excel workbook and confirm they stay in the original language after localization with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file, accesses the first worksheet's first chart, and enables series name, category name, and value data labels for every series using Aspose.Cells. | Show how to programmatically verify that chart data labels retain their original language after the workbook has been localized with Aspose.Cells. | Provide a snippet that adds data labels to a chart, saves the workbook to a new file, and logs a message confirming the label text was not altered by localization.
// Common Searches: aspocells add series name and value data labels to chart after localization | C# verify Excel chart data labels keep original language using Aspose.Cells | how to enable category name data labels in Aspose.Cells chart | preserve data label language when localizing workbook with Aspose.Cells .NET | add data labels to first chart in workbook Aspose.Cells example
// Tags: configure chart data labels Aspose.Cells .NET | display series name category name value in Excel chart C# | check data label language after workbook localization Aspose.Cells | preserve original language in chart data labels .NET | enable data labels for all series using Aspose.Cells API

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// // Loads an existing Excel file, accesses the first worksheet's first chart, enables series name, category name, and value data labels for each series, saves the modified workbook, and can be used to confirm that labels remain in the original language after localization.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Add data labels to each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Configure which parts of the label should be shown
                series.DataLabels.ShowSeriesName = true;      // Show series name
                series.DataLabels.ShowCategoryName = true;    // Show category (X‑axis) name
                series.DataLabels.ShowValue = true;           // Show the actual value
            }

            // Save the workbook with the newly added data labels
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
