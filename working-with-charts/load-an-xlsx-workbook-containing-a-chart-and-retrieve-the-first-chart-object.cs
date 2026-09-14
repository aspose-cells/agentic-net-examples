// Title: How to load an XLSX workbook and retrieve the first chart object using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, verifies that the first worksheet contains charts, and returns the Chart object at zero‑based index 0. | Write a C# console application that loads a workbook, checks the chart count on the first worksheet, extracts the initial chart, and prints its Name, Type, and parent worksheet.
// Common Searches: Aspose.Cells C# get chart from first worksheet of an existing XLSX file | C# retrieve chart object index 0 using Aspose.Cells Workbook | How to list all charts in a worksheet with Aspose.Cells .NET API | Read chart properties (name, type) from an XLSX workbook using Aspose.Cells C#
// Tags: load xlsx workbook with Aspose.Cells C# | enumerate worksheet charts Aspose.Cells | extract chart by zero-based index Aspose.Cells | display chart name and type Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartRetrieval
{
    // The example demonstrates loading an existing XLSX file with Aspose.Cells, accessing the first worksheet, confirming the presence of charts, retrieving the first chart object, and outputting its name, type, and the worksheet it belongs to.
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX file that contains at least one chart
            string filePath = "input.xlsx";

            // Load the workbook from the file (uses the Workbook(string) constructor rule)
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count > 0)
            {
                // Retrieve the first chart object (index 0) from the Charts collection
                Chart firstChart = worksheet.Charts[0];

                // Example usage: display chart properties
                Console.WriteLine($"Chart Name: {firstChart.Name}");
                Console.WriteLine($"Chart Type: {firstChart.Type}");
                Console.WriteLine($"Containing Worksheet: {firstChart.Worksheet.Name}");
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }
        }
    }
}
