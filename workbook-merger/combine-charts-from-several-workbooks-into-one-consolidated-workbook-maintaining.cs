using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartConsolidation
{
    public class ConsolidateCharts
    {
        public static void Run()
        {
            // Paths of the source workbooks that contain charts
            string[] sourceFiles = new string[]
            {
                "ChartWorkbook1.xlsx",
                "ChartWorkbook2.xlsx",
                "ChartWorkbook3.xlsx"
            };

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Keep references to the loaded source workbooks for later link updates
            List<Workbook> sourceWorkbooks = new List<Workbook>();

            // Load each source workbook, combine its content (including charts) into the destination,
            // and store the workbook reference for possible external link refresh
            foreach (string filePath in sourceFiles)
            {
                Workbook source;

                if (File.Exists(filePath))
                {
                    source = new Workbook(filePath);
                }
                else
                {
                    // Create a dummy workbook with a simple chart if the file does not exist
                    source = new Workbook();
                    Worksheet ws = source.Worksheets[0];
                    ws.Name = "Data";

                    // Populate sample data
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Value");
                    ws.Cells["A2"].PutValue("A");
                    ws.Cells["B2"].PutValue(10);
                    ws.Cells["A3"].PutValue("B");
                    ws.Cells["B3"].PutValue(20);
                    ws.Cells["A4"].PutValue("C");
                    ws.Cells["B4"].PutValue(30);

                    // Add a chart
                    int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                    Chart chart = ws.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";
                    chart.Title.Text = "Sample Chart";

                    // Save the dummy workbook for future runs (optional)
                    source.Save(filePath, SaveFormat.Xlsx);
                }

                sourceWorkbooks.Add(source);
                destinationWorkbook.Combine(source);
            }

            // Refresh all charts (and pivot tables) to ensure they reflect the combined data.
            destinationWorkbook.Worksheets.RefreshAll();

            // Update external links using the source workbooks, if any.
            destinationWorkbook.UpdateLinkedDataSource(sourceWorkbooks.ToArray());

            // Save the consolidated workbook with all charts, data links, and formatting preserved.
            destinationWorkbook.Save("ConsolidatedCharts.xlsx", SaveFormat.Xlsx);

            // Dispose all workbooks to release resources.
            foreach (Workbook wb in sourceWorkbooks)
            {
                wb.Dispose();
            }
            destinationWorkbook.Dispose();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConsolidateCharts.Run();
        }
    }
}