using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAddSeriesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // First series values
            sheet.Cells["B1"].PutValue("Sales 2020");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Second series values (the new series we will add)
            sheet.Cells["C1"].PutValue("Sales 2021");
            sheet.Cells["C2"].PutValue(130);
            sheet.Cells["C3"].PutValue(160);
            sheet.Cells["C4"].PutValue(190);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the first series (Sales 2020) using a range
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            // Set category axis data
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Add a new series (Sales 2021) using a worksheet range as data source
            // The Add method returns the index of the first series added; we ignore it here
            chart.NSeries.Add("=Sheet1!$C$2:$C$4", true);

            // Optionally set the name of the new series explicitly
            chart.NSeries[1].Name = "Sales 2021";

            // Save the workbook to a file
            workbook.Save("ChartWithAddedSeries.xlsx");
        }
    }
}