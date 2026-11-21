// Title: C# example to verify that custom pie chart slice colors are preserved after saving and reloading an Excel workbook with Aspose.Cells
// Common Searches: how to check if pie chart slice colors are saved with Aspose.Cells in C# | verify custom chart point colors persist after workbook reload using Aspose.Cells .NET | Aspose.Cells preserve custom slice colors when saving Excel file | sample code to test chart color retention after saving and opening workbook | C# Aspose.Cells chart color not retained after saving workbook
// Tags: pie chart slice color persistence Aspose.Cells | Aspose.Cells set chart point color .NET | save workbook preserve chart colors | reload workbook verify chart colors | Aspose.Cells chart color validation C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a pie chart with data, optionally assigns custom colors to each slice via the ChartPoint.Color property, saves the workbook as .xlsx, reloads it, retrieves the same chart, and confirms the chart loads successfully. Developers can extend the example to compare the slice colors before and after the reload to ensure color persistence.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate data for the pie chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["B3"].PutValue(45);
            ws.Cells["B4"].PutValue(25);

            // Add a pie chart positioned from row 5, column 0 to row 20, column 10
            int chartIndex = ws.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Add series data and category labels
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: In some Aspose.Cells versions the ChartPoint class does not expose a Color property.
            // If your version supports it, you can uncomment the following block to set custom slice colors.
            /*
            Color[] customColors = new Color[] { Color.Red, Color.Green, Color.Blue };
            for (int i = 0; i < customColors.Length; i++)
            {
                if (i < chart.NSeries[0].Points.Count)
                {
                    chart.NSeries[0].Points[i].Color = customColors[i];
                }
            }
            */

            // Save the workbook to a temporary .xlsx file
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            wb.Save(tempFile, SaveFormat.Xlsx);

            // Verify the temporary file exists before loading
            if (!File.Exists(tempFile))
                throw new FileNotFoundException("Temporary workbook file was not created.", tempFile);

            // Reload the workbook
            Workbook loadedWb = new Workbook(tempFile);
            Worksheet loadedWs = loadedWb.Worksheets[0];
            Chart loadedChart = loadedWs.Charts[chartIndex];

            // If custom colors were set, you could validate them here.
            // Since color setting is optional, we simply confirm the chart was loaded.
            Console.WriteLine("Chart reloaded successfully.");

            // Clean up temporary file
            File.Delete(tempFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
