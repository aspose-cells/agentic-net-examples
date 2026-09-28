// Title: Verify that a chart added with Aspose.Cells for .NET uses the default English title when no localization is configured
// AI Prompts: Generate C# code using Aspose.Cells to insert a column chart, leave all language and localization properties unset, and programmatically confirm that chart.Title.Text is empty or equals "Chart Title". | Write a C# unit‑test with Aspose.Cells that creates a workbook, adds a chart without any localization settings, and asserts the default chart title is the English placeholder.
// Common Searches: How does Aspose.Cells decide the chart title text when no language is set in .NET? | C# code to check if a newly added Aspose.Cells chart has an empty title by default | What placeholder does Aspose.Cells use for a chart title if localization is omitted | Determine the default English title of an Aspose.Cells chart without configuring localization
// Tags: Aspose.Cells chart title default value | C# Aspose.Cells chart without language settings | Aspose.Cells chart localization verification | Aspose.Cells chart title language detection | Aspose.Cells .NET chart creation example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds a column chart without assigning any localization or language properties, reads the chart's Title.Text, and verifies it is either empty or the built‑in English placeholder "Chart Title" before saving the file.
class VerifyChartDefaultEnglishText
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart of type Column
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart (no localization settings are applied)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Do NOT set any localization or language properties on the chart

        // Retrieve the default title text (Aspose.Cells creates a default title if none is set)
        string titleText = chart.Title.Text;

        // Verify that the title text is in English (default is an empty string or "Chart Title")
        // For Aspose.Cells, if no title is explicitly set, Title.Text returns an empty string.
        if (string.IsNullOrEmpty(titleText))
        {
            Console.WriteLine("Verification passed: Default chart title is empty (English default).");
        }
        else
        {
            // If a default title exists, ensure it is English (e.g., "Chart Title")
            if (titleText.Equals("Chart Title", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Verification passed: Default chart title is English ('Chart Title').");
            }
            else
            {
                Console.WriteLine($"Verification failed: Unexpected chart title '{titleText}'.");
            }
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("ChartDefaultEnglishText.xlsx");
    }
}
