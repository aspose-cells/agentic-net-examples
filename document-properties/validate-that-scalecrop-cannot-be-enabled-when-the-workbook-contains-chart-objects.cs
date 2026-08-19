// Title: Aspose.Cells .NET – Disable ScaleCrop When Workbook Contains Charts
// Description: C# example that creates a workbook, adds a column chart, scans all worksheets for chart objects, and conditionally sets BuiltInDocumentProperties.ScaleCrop to false before saving.
// Keywords: Aspose.Cells | .NET | C# | ScaleCrop | chart detection | built‑in document properties | Excel workbook validation | disable ScaleCrop | chart objects | SaveFormat.Xlsx
// Common Searches: Aspose.Cells disable ScaleCrop with charts | ScaleCrop property restriction chart objects | check for charts before setting ScaleCrop in .NET | how to validate ScaleCrop in Excel using Aspose.Cells | C# code to prevent ScaleCrop when workbook has charts
// Developer Intent: Programmatically ensure ScaleCrop is turned off whenever any worksheet in the workbook contains a chart.
// Use Cases: Automatically turn off ScaleCrop in generated reports that include charts to avoid rendering problems. | Validate workbook content before export in a CI/CD pipeline that creates Excel files with optional visualizations. | Implement a lifecycle rule that conditionally applies ScaleCrop based on the presence of chart objects.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets, detects charts, and disables the ScaleCrop property if any are found. | Create a reusable method that accepts a Workbook and returns true if ScaleCrop can be safely enabled, considering chart objects. | Show how to wrap ScaleCrop validation in a try‑catch block while exporting a workbook that may contain charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Properties;

// C# example that creates a workbook, adds a column chart, scans all worksheets for chart objects, and conditionally sets BuiltInDocumentProperties.ScaleCrop to false before saving.
class ScaleCropValidationDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

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

        // Validate ScaleCrop setting according to presence of charts
        if (containsChart)
        {
            // ScaleCrop must not be enabled when charts exist
            properties.ScaleCrop = false;
            Console.WriteLine("ScaleCrop disabled because the workbook contains chart objects.");
        }
        else
        {
            properties.ScaleCrop = true;
            Console.WriteLine("ScaleCrop enabled.");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("ScaleCropValidation.xlsx", SaveFormat.Xlsx);
    }
}
