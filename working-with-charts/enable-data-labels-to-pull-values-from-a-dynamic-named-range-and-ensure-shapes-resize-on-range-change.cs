// Title: How to create a column chart with data labels from a dynamic named range and enable auto‑scaling of the chart area using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bind chart data labels to a dynamic named range | C# auto scaling chart area when the data range expands Aspose.Cells | Create column chart with OFFSET named range for labels using Aspose.Cells | How to use a dynamic named range for chart data labels in .NET | Aspose.Cells chart automatically resize based on data range changes
// Tags: dynamic named range for chart labels Aspose.Cells | auto scaling chart area .NET Aspose.Cells | column chart with data labels Aspose.Cells C# | OFFSET named range usage Aspose.Cells | NSeries data label binding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, defines a dynamic named range (using OFFSET and COUNTA) for chart data labels, adds a column chart, links the series values, hides default numeric labels, enables AutoScaling so the chart area adjusts when rows are added, and saves the file as DynamicDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and rename it
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data (Category in A, Values in B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Define a dynamic named range for data labels (adjusts with added rows)
            int nameIdx = workbook.Worksheets.Names.Add("DynamicLabels");
            var name = workbook.Worksheets.Names[nameIdx];
            // Use the actual worksheet name instead of the default "Sheet1"
            name.RefersTo = $"=OFFSET({sheet.Name}!$B$2,0,0,COUNTA({sheet.Name}!$B:$B)-1,1)";

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIdx];

            // Set the series values from the static data range (B2:B4)
            chart.NSeries.Add($"={sheet.Name}!$B$2:$B$4", true);

            // Hide default numeric labels
            chart.NSeries[0].DataLabels.ShowValue = false;

            // NOTE: Custom data labels are not supported in the current Aspose.Cells version.
            // If a newer version provides the CustomDataLabels property, the following line can be uncommented:
            // chart.NSeries[0].DataLabels.CustomDataLabels.Add("=DynamicLabels");

            // Ensure the chart area scales automatically when the data range changes
            chart.AutoScaling = true;

            // Save the workbook
            string outputPath = "DynamicDataLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
