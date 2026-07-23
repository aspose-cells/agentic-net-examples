// Title: C# – Disable ScaleCrop When Workbook Contains Charts Using Aspose.Cells
// Description: Shows how to create a workbook, add a column chart, scan all worksheets for chart objects, and set the built‑in ScaleCrop property to false (or true when no charts) before saving the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | ScaleCrop | C# chart detection | built‑in document properties | disable ScaleCrop | Excel chart check | Aspose.Cells .NET | XLSX export | global developers | US .NET community
// Common Searches: Aspose.Cells set ScaleCrop based on chart presence | C# detect charts in workbook before saving | disable ScaleCrop when Excel file has charts | built‑in document property ScaleCrop Aspose.Cells | how to turn off ScaleCrop for chart‑filled workbook
// Developer Intent: Automatically set ScaleCrop to false whenever any worksheet in the workbook contains a chart.
// Use Cases: Prevent image‑scaling artifacts by disabling ScaleCrop in workbooks that include visualizations. | Enforce corporate document‑property policies programmatically before exporting Excel files. | Batch‑process generated reports, toggling ScaleCrop according to the presence of charts.
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, detects chart objects, and sets BuiltInDocumentPropertyCollection.ScaleCrop to false if any are found. | Provide an example that toggles ScaleCrop based on chart detection and saves the workbook as an XLSX file. | Explain how to unit‑test the ScaleCrop validation logic for workbooks with and without charts using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Properties;

// Shows how to create a workbook, add a column chart, scan all worksheets for chart objects, and set the built‑in ScaleCrop property to false (or true when no charts) before saving the file with Aspose.Cells for .NET.
class ScaleCropValidation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for a chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);

        // Add a chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Check if any worksheet contains chart objects
        bool containsChart = false;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Charts.Count > 0)
            {
                containsChart = true;
                break;
            }
        }

        // Access built‑in document properties
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Validate ScaleCrop: it must be false when charts are present
        if (containsChart)
        {
            properties.ScaleCrop = false;
            Console.WriteLine("Workbook contains chart(s); ScaleCrop disabled.");
        }
        else
        {
            properties.ScaleCrop = true;
            Console.WriteLine("No charts found; ScaleCrop enabled.");
        }

        // Save the workbook
        workbook.Save("ScaleCropValidation.xlsx", SaveFormat.Xlsx);
    }
}
