// Title: Aspose.Cells for .NET: Add a Pie Chart with Percentage Labels and Auto‑Fit Data‑Label Shape (C#)
// Description: Creates a workbook, fills cells with categories and values, inserts a pie chart, binds the series to the value range, enables data labels to show only percentages, hides raw values, activates auto‑fit for the label shape with an initial width, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | pie chart | percentage data labels | auto‑fit label shape | IsResizeShapeToFitText | ShowPercentage | Excel chart customization | data label width | chart export
// Common Searches: Aspose.Cells show percentage on pie chart | C# set data label auto size Aspose.Cells | hide values display only percentages Aspose.Cells chart | resize data label shape to fit text Aspose.Cells
// Developer Intent: Generate a pie chart where data labels display only the percentage values from worksheet cells and the label shapes automatically resize to fit the text.
// Use Cases: Sales distribution report with clean percentage‑only slices | Market‑share dashboard where label size adapts to varying percentages | Executive‑level workbook export with auto‑fitting pie‑chart labels
// AI Prompts: Write C# code using Aspose.Cells to add a pie chart, bind series to B2:B4, show only percentages in data labels, enable IsResizeShapeToFitText, set an initial WidthPixel, hide raw values, and save the workbook. | Explain step‑by‑step how to configure percentage‑only data labels and auto‑fit label shapes in an Aspose.Cells pie chart (C#). | Provide a minimal reproducible example that demonstrates adjusting the data label shape size based on percentage text length in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills cells with categories and values, inserts a pie chart, binds the series to the value range, enables data labels to show only percentages, hides raw values, activates auto‑fit for the label shape with an initial width, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate worksheet with categories and values
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and configure them to show percentages from cells
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowPercentage = true;   // display percentage
        dataLabels.ShowValue = false;       // hide raw value
        dataLabels.ShowCategoryName = false;

        // Adjust shape size: enable auto‑fit to text and set an initial width
        dataLabels.IsResizeShapeToFitText = true;
        dataLabels.WidthPixel = 80; // base width; will expand if needed to fit the percentage text

        // Save the workbook
        workbook.Save("PieChartWithPercentLabels.xlsx");
    }
}
