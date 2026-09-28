// Title: Create a line chart with a DateAxis (TimeScale) and monthly base units using Aspose.Cells for .NET
// AI Prompts: Generate a C# workbook that adds a line chart, assigns date values to the X‑axis, and sets the category axis to TimeScale with months as the base unit. | Modify an existing Aspose.Cells chart to display dates on the category axis by changing CategoryType to TimeScale and configuring BaseUnitScale to months. | Produce sample code that saves an XLSX file containing a line chart where the X‑axis is a DateAxis formatted for monthly intervals.
// Common Searches: aspnet how to set chart category axis to dates using Aspose.Cells | Aspose.Cells line chart time series X axis month scale | C# change chart CategoryType to TimeScale in Excel workbook | set BaseUnitScale months for date axis Aspose.Cells chart | display monthly date axis in Aspose.Cells line chart
// Tags: Aspose.Cells set CategoryAxis to TimeScale | Aspose.Cells line chart DateAxis months | C# chart category type DateAxis example | Aspose.Cells base unit scale months | Aspose.Cells time series chart configuration

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills column A with DateTime values and column B with numeric data, adds a line chart, links the series to the data ranges, changes the category axis to a TimeScale (DateAxis), sets the base unit scale to months, and saves the file as DateAxisChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with date‑based data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X) axis
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Change the category axis type to DateAxis (TimeScale) for proper time‑based display
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: define the base unit scale (e.g., months) for the date axis
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;

        // Save the workbook
        workbook.Save("DateAxisChart.xlsx", SaveFormat.Xlsx);
    }
}
