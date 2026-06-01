using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Properties;

class ValidateScaleCrop
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for a chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Determine whether the workbook contains any chart objects
        bool hasChart = false;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Charts.Count > 0)
            {
                hasChart = true;
                break;
            }
        }

        // Access built‑in document properties
        BuiltInDocumentPropertyCollection props = workbook.BuiltInDocumentProperties;

        if (hasChart)
        {
            // ScaleCrop must not be enabled when charts exist
            Console.WriteLine("Workbook contains chart objects; ScaleCrop will remain disabled.");
            props.ScaleCrop = false;
        }
        else
        {
            // No charts present, safe to enable ScaleCrop
            props.ScaleCrop = true;
            Console.WriteLine("ScaleCrop enabled: " + props.ScaleCrop);
        }

        // Save the workbook
        workbook.Save("ValidateScaleCrop.xlsx", SaveFormat.Xlsx);
    }
}