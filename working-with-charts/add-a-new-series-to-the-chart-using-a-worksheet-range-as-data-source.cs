using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddSeriesFromRange
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Category column
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        // First series values
        sheet.Cells["B1"].PutValue("Sales 2022");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(130);
        sheet.Cells["B5"].PutValue(170);

        // Second series values (the new series we will add)
        sheet.Cells["C1"].PutValue("Sales 2023");
        sheet.Cells["C2"].PutValue(140);
        sheet.Cells["C3"].PutValue(160);
        sheet.Cells["C4"].PutValue(150);
        sheet.Cells["C5"].PutValue(180);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the category (X‑axis) data for the chart
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

        // Add the first series (2022) using a range
        chart.NSeries.Add("=Sheet1!$B$2:$B$5", true);

        // Add the new series (2023) using a worksheet range as data source
        // The Add method takes the data range string and a boolean indicating vertical orientation
        chart.NSeries.Add("=Sheet1!$C$2:$C$5", true);

        // Optional: give each series a name (if not already set by the first row)
        // Here we set names explicitly using SetSeriesNames starting at index 0
        chart.NSeries.SetSeriesNames(0, "B1:C1", true);

        // Save the workbook to a file
        workbook.Save("ChartWithAdditionalSeries.xlsx");
    }
}