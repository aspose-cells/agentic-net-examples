using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ListCharts
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Add sample data and a chart to the first worksheet
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Cells["A1"].PutValue("Category");
            sheet1.Cells["A2"].PutValue("A");
            sheet1.Cells["A3"].PutValue("B");
            sheet1.Cells["A4"].PutValue("C");
            sheet1.Cells["B1"].PutValue("Value");
            sheet1.Cells["B2"].PutValue(10);
            sheet1.Cells["B3"].PutValue(20);
            sheet1.Cells["B4"].PutValue(30);

            int chartIdx1 = sheet1.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart1 = sheet1.Charts[chartIdx1];
            chart1.Name = "SalesChart";
            chart1.NSeries.Add("B2:B4", true);
            chart1.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Add a second worksheet with a pie chart
            // -------------------------------------------------
            workbook.Worksheets.Add("SecondSheet");               // add sheet by name
            Worksheet sheet2 = workbook.Worksheets["SecondSheet"]; // retrieve by name
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["A2"].PutValue("X");
            sheet2.Cells["A3"].PutValue("Y");
            sheet2.Cells["B1"].PutValue("Qty");
            sheet2.Cells["B2"].PutValue(5);
            sheet2.Cells["B3"].PutValue(15);

            int chartIdx2 = sheet2.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
            Chart chart2 = sheet2.Charts[chartIdx2];
            chart2.Name = "DistributionChart";
            chart2.NSeries.Add("B2:B3", true);
            chart2.NSeries.CategoryData = "A2:A3";

            // -------------------------------------------------
            // Iterate through all worksheets and list each chart's name and type
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {ws.Name}");
                ChartCollection charts = ws.Charts;
                for (int i = 0; i < charts.Count; i++)
                {
                    Chart c = charts[i];
                    string name = string.IsNullOrEmpty(c.Name) ? "(no name)" : c.Name;
                    Console.WriteLine($"  Chart {i}: Name = {name}, Type = {c.Type}");
                }
            }

            // Save the workbook (ensure the directory exists)
            string outputPath = "ListChartsOutput.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}