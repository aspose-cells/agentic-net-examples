using Aspose.Cells;
using Aspose.Cells.Charts;

class SetXAxisTitle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B4"].PutValue(1800);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the X‑axis (CategoryAxis) title to a descriptive label
        chart.CategoryAxis.Title.Text = "Months (Jan‑Mar)";
        chart.CategoryAxis.Title.IsVisible = true;

        // Optional: set a chart title for completeness
        chart.Title.Text = "Quarterly Sales";
        chart.Title.IsVisible = true;

        // Save the workbook to a file
        workbook.Save("XAxisTitleDemo.xlsx");
    }
}