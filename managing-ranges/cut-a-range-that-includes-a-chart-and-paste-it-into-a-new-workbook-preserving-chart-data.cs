using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

class CutRangeWithChart
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook ----------
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];
            srcSheet.Name = "Source";

            // Fill sample data (A1:B5)
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                srcSheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                srcSheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a chart that uses the data range A1:B5
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 7, 0, 20, 5);
            Chart srcChart = srcSheet.Charts[chartIdx];
            srcChart.NSeries.Add("=Source!$B$2:$B$5", true);
            srcChart.NSeries.CategoryData = "=Source!$A$2:$A$5";
            srcChart.Title.Text = "Sample Chart";

            // ---------- Define the range to cut (including data) ----------
            // The chart itself is a drawing object, not part of cells, so we only cut the data range.
            AsposeRange srcDataRange = srcSheet.Cells.CreateRange("A1:B5");

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];
            destSheet.Name = "Destination";

            // Destination range where data will be pasted (same size)
            AsposeRange destDataRange = destSheet.Cells.CreateRange("A1:B5");

            // ---------- Copy (cut) the data range ----------
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All // copy values, formulas, formats, etc.
            };
            destDataRange.Copy(srcDataRange, pasteOptions);

            // ---------- Recreate the chart in the destination sheet ----------
            // Use the same chart type and position; adjust data source to the destination sheet.
            int destChartIdx = destSheet.Charts.Add(srcChart.Type, 7, 0, 20, 5);
            Chart destChart = destSheet.Charts[destChartIdx];
            // Data source now points to the destination sheet.
            destChart.NSeries.Add("=Destination!$B$2:$B$5", true);
            destChart.NSeries.CategoryData = "=Destination!$A$2:$A$5";
            destChart.Title.Text = srcChart.Title.Text;

            // ---------- Save both workbooks ----------
            string srcPath = "SourceWorkbook.xlsx";
            string destPath = "DestinationWorkbook.xlsx";

            // Ensure we can write to the target locations
            try
            {
                srcWb.Save(srcPath);
                destWb.Save(destPath);
                Console.WriteLine($"Workbooks saved successfully: {srcPath}, {destPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbooks: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions to prevent crashes
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}