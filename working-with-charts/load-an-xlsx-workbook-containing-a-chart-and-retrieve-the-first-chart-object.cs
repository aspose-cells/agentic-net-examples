// Title: Load an XLSX workbook and get the first chart with Aspose.Cells for .NET
// Description: Shows how to open an existing XLSX file using Aspose.Cells for .NET, access the first worksheet, verify its Charts collection, retrieve the first Chart object, and read its Name and Type properties.
// Keywords: Aspose.Cells load workbook | Aspose.Cells chart collection | retrieve first chart C# | read chart name Aspose.Cells | Excel chart object .NET | Aspose.Cells get chart | Workbook(string) constructor | chart properties Aspose.Cells | C# Excel chart extraction | Aspose.Cells example
// Common Searches: How to read the first chart in an Excel file using Aspose.Cells for .NET | Aspose.Cells retrieve chart object from worksheet | Get chart name and type from XLSX with Aspose.Cells | Check if a worksheet contains charts using Aspose.Cells | Aspose.Cells C# load workbook and access charts
// Developer Intent: Obtain the first Chart object from the first worksheet of an existing XLSX workbook using Aspose.Cells for .NET.
// Use Cases: Display the chart's name and type to confirm successful retrieval. | Validate the presence of charts before applying chart‑specific logic. | Modify chart attributes (e.g., title, style) after acquiring the Chart instance.
// AI Prompts: Provide C# code that loads an XLSX file with Aspose.Cells and extracts the data series of the first chart on the first worksheet. | Show how to change the title of the first chart after retrieving it from a worksheet using Aspose.Cells for .NET. | Explain how to loop through all charts in a workbook and export each chart as an image with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to open an existing XLSX file using Aspose.Cells for .NET, access the first worksheet, verify its Charts collection, retrieve the first Chart object, and read its Name and Type properties.
class Program
{
    static void Main()
    {
        // Load the existing workbook that contains a chart
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // uses Workbook(string) constructor

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the first chart object from the worksheet's chart collection
        if (worksheet.Charts.Count > 0)
        {
            Chart firstChart = worksheet.Charts[0]; // ChartCollection indexer

            // Example: display some properties of the retrieved chart
            Console.WriteLine($"Chart Name: {firstChart.Name}");
            Console.WriteLine($"Chart Type: {firstChart.Type}");
        }
        else
        {
            Console.WriteLine("No charts found in the first worksheet.");
        }

        // (Optional) Save the workbook if any modifications are made
        // workbook.Save("output.xlsx");
    }
}
