// Title: Add a column chart with a dynamic secondary‑axis data range using INDEX/MATCH formulas in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, populates sample data, adds a column chart, and defines the secondary series range with an INDEX/MATCH formula. | Demonstrate how to assign a chart series to the secondary axis and use a dynamic range based on lookup values in Aspose.Cells.
// Common Searches: how to use INDEX and MATCH in Aspose.Cells to define a chart series range | Aspose.Cells C# create column chart with secondary axis and dynamic data range | programmatically set secondary axis series source using Excel formulas in .NET | dynamic range for secondary axis chart using Aspose.Cells chart API | C# example of chart.NSeries.Add with INDEX/MATCH formula in Aspose.Cells
// Tags: add column chart secondary axis Aspose.Cells | set chart series data range with INDEX MATCH | dynamic secondary axis range C# | chart series formula Aspose.Cells | assign series to secondary axis .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Shows how to create a workbook, fill it with categories and two data series, add a column chart, define the primary series with a static range, use an INDEX/MATCH formula to build a dynamic range for the secondary series, plot that series on the secondary axis, and save the file as an .xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            // Column A: Categories (e.g., months)
            // Column B: Series1 values (primary axis)
            // Column C: Series2 values (secondary axis)
            // Column D: Dummy lookup column (not used directly in chart)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Lookup");

            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] series1 = { 10, 20, 30, 40, 50 };
            double[] series2 = { 15, 25, 35, 45, 55 };
            string[] lookup = { "Apr", "May" }; // Example lookup values

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]); // A column
                sheet.Cells[i + 1, 1].PutValue(series1[i]);   // B column
                sheet.Cells[i + 1, 2].PutValue(series2[i]);   // C column
                sheet.Cells[i + 1, 3].PutValue(lookup[i % lookup.Length]); // D column (optional)
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Dynamic Secondary Axis";

            // Primary series (Series1) using a static range
            chart.NSeries.Add("=Data!$B$2:$B$6", true);
            chart.NSeries[0].Name = "Series1";

            // Secondary series (Series2) using INDEX/MATCH to create a dynamic range
            // Formula: =INDEX(Data!$C:$C, MATCH("Apr",Data!$A:$A,0)):INDEX(Data!$C:$C, MATCH("May",Data!$A:$A,0))
            string startMatch = "MATCH(\"Apr\",Data!$A:$A,0)";
            string endMatch = "MATCH(\"May\",Data!$A:$A,0)";
            string seriesFormula = $"=INDEX(Data!$C:$C,{startMatch}):INDEX(Data!$C:$C,{endMatch})";

            // Add the secondary series with the dynamic formula
            chart.NSeries.Add(seriesFormula, true);
            chart.NSeries[1].Name = "Series2";

            // Assign the secondary series to the secondary axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Save the workbook
            workbook.Save("DynamicSecondaryAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
