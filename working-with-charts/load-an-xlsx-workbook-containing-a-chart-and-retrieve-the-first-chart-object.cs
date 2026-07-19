// Title: Load an XLSX workbook and retrieve the first chart using Aspose.Cells for .NET
// Description: Demonstrates how to open an existing XLSX file with Aspose.Cells, access the first worksheet, check for charts, and obtain the first Chart object to read its Name and Type.
// Keywords: Aspose.Cells load workbook | C# read chart from XLSX | Aspose.Cells first chart | chart collection Aspose.Cells | retrieve chart name type | .NET spreadsheet chart API
// Common Searches: Aspose.Cells get first chart from worksheet | C# load XLSX and read chart properties | How to check chart count in Aspose.Cells | Retrieve chart name and type using Aspose.Cells | Access chart collection in Aspose.Cells .NET
// Developer Intent: Load an existing XLSX file and obtain the first chart object from the first worksheet for further processing.
// Use Cases: Display or log the name and type of the first chart after opening a workbook. | Validate chart presence before performing export or manipulation to avoid runtime errors. | Iterate over a worksheet's chart collection to extract properties or generate images.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, checks for charts, and returns the first Chart object safely. | Create a method that accepts a Worksheet and returns the first Chart, including a null‑check for an empty chart collection. | Show how to export the first chart to a PNG file using Aspose.Cells after retrieving it.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to open an existing XLSX file with Aspose.Cells, access the first worksheet, check for charts, and obtain the first Chart object to read its Name and Type.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook that contains a chart
        string inputPath = "input.xlsx";               // path to the source file
        Workbook workbook = new Workbook(inputPath);    // uses Workbook(string) constructor

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the first chart object from the worksheet's chart collection
        if (worksheet.Charts.Count > 0)
        {
            Chart firstChart = worksheet.Charts[0];    // ChartCollection indexer (zero‑based)

            // Example usage: display chart information
            Console.WriteLine($"First chart name: {firstChart.Name}");
            Console.WriteLine($"First chart type: {firstChart.Type}");
        }
        else
        {
            Console.WriteLine("No charts found in the first worksheet.");
        }
    }
}
