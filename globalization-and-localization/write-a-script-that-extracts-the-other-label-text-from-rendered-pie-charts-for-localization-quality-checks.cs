// Title: C# – Retrieve Localized “Other” Segment Label from a Pie Chart with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a pie chart, force chart calculation, and use ChartGlobalizationSettings.GetOtherName() to obtain the localized text for the “Other” segment label, output it, and save the file.
// Keywords: Aspose.Cells C# pie chart localization | ChartGlobalizationSettings GetOtherName | extract localized chart label .NET | retrieve "Other" segment text | globalization settings Aspose.Cells | localization quality check workbook | C# Aspose.Cells chart label extraction | pie chart other label translation | Aspose.Cells localization testing
// Common Searches: Aspose.Cells get localized "Other" label from pie chart | ChartGlobalizationSettings GetOtherName example C# | how to read internal pie chart labels Aspose.Cells | C# extract pie chart segment name for localization | Aspose.Cells chart globalization retrieve other name
// Developer Intent: The developer needs to programmatically obtain the localized string used for the “Other” segment label in a pie chart created with Aspose.Cells.
// Use Cases: Verify that the "Other" label matches expected translations across multiple cultures before releasing a workbook. | Generate a QA report that lists the localized "Other" labels from all charts in a set of workbooks. | Automate regression tests that compare the "Other" label after changing the culture in ChartGlobalizationSettings.
// AI Prompts: Show how to set a specific culture on ChartGlobalizationSettings and then retrieve the "Other" label in that language using C#. | Provide code that loops through every chart in a workbook and prints each chart's localized "Other" label. | Explain how to validate that the "Other" label is correctly localized after applying custom globalization settings in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings; // Namespace for ChartGlobalizationSettings

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a pie chart, force chart calculation, and use ChartGlobalizationSettings.GetOtherName() to obtain the localized text for the “Other” segment label, output it, and save the file.
    public class ExtractOtherLabelFromPieChart
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["A5"].PutValue("Date");
            sheet.Cells["A6"].PutValue("Elderberry");
            // Value column
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["B5"].PutValue(15);
            sheet.Cells["B6"].PutValue(10);

            // Add a pie chart that will render the data
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 8);
            Chart pieChart = sheet.Charts[chartIndex];
            pieChart.NSeries.Add("B2:B6", true);          // Values
            pieChart.NSeries.CategoryData = "A2:A6";     // Categories

            // Force chart calculation so that internal labels are generated
            pieChart.Calculate();

            // Retrieve the localized name for the "Other" label using globalization settings
            ChartGlobalizationSettings globalization = new ChartGlobalizationSettings();
            string otherLabel = globalization.GetOtherName();

            // Output the retrieved "Other" label text
            Console.WriteLine("Localized 'Other' label text: " + otherLabel);

            // Save the workbook (the chart is saved as part of the workbook)
            workbook.Save("PieChartOtherLabelDemo.xlsx");
        }
    }
}
