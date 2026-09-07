// Title: Load an Excel workbook from a MemoryStream and enumerate its chart collection with Aspose.Cells for .NET
// AI Prompts: Create a Workbook from a MemoryStream and output each chart’s type, title, and series count using Aspose.Cells in C#. | Read an XLS file into a MemoryStream, instantiate a Workbook, and loop through the worksheet’s ChartCollection to process chart metadata.
// Common Searches: Aspose.Cells C# open Excel file from a MemoryStream and list charts | How to get chart collection from a workbook loaded via stream in .NET | Enumerate Excel charts after loading workbook from a byte array using Aspose.Cells
// Tags: load workbook via stream Aspose.Cells | enumerate worksheet chart collection C# | extract chart type and title Aspose.Cells | process Excel charts from streamed workbook | read XLS from MemoryStream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example builds a workbook with sample data and a column chart, saves it to a MemoryStream, loads a new Workbook from that stream, and then iterates the first worksheet's ChartCollection, printing each chart's type, title, and series count.
    public class LoadWorkbookFromMemoryStream
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook with a chart (optional, for demo)
            // ------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Add sample data
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["A4"].PutValue("C");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a chart to the worksheet
            int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart sourceChart = sourceSheet.Charts[chartIdx];
            sourceChart.NSeries.Add("B2:B4", true);
            sourceChart.NSeries.CategoryData = "A2:A4";
            sourceChart.Title.Text = "Sample Chart";

            // ------------------------------------------------------------
            // 2. Save the workbook to a MemoryStream (as XLS)
            // ------------------------------------------------------------
            using (MemoryStream memoryStream = sourceWorkbook.SaveToStream())
            {
                // Reset the stream position before reading
                memoryStream.Position = 0;

                // ------------------------------------------------------------
                // 3. Load a new workbook from the MemoryStream
                // ------------------------------------------------------------
                using (Workbook loadedWorkbook = new Workbook(memoryStream))
                {
                    // ------------------------------------------------------------
                    // 4. Access the chart collection of the first worksheet
                    // ------------------------------------------------------------
                    Worksheet firstSheet = loadedWorkbook.Worksheets[0];
                    ChartCollection charts = firstSheet.Charts;

                    // Example processing: iterate through all charts and display basic info
                    for (int i = 0; i < charts.Count; i++)
                    {
                        Chart chart = charts[i];
                        Console.WriteLine($"Chart #{i + 1}");
                        Console.WriteLine($"  Type : {chart.Type}");
                        Console.WriteLine($"  Title: {chart.Title.Text}");
                        Console.WriteLine($"  Series count: {chart.NSeries.Count}");
                    }
                }
            }

            // ------------------------------------------------------------
            // 5. Clean up resources
            // ------------------------------------------------------------
            sourceWorkbook.Dispose();
        }
    }
}
