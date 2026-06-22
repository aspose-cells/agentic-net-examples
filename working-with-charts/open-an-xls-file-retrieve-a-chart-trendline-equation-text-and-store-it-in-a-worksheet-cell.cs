using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RetrieveTrendlineEquation
{
    static void Main()
    {
        // Load the existing XLS file
        string inputPath = "input.xls";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart in the worksheet
        Chart chart = worksheet.Charts[0];

        // Ensure the first series has at least one trendline
        if (chart.NSeries.Count == 0 || chart.NSeries[0].TrendLines.Count == 0)
        {
            Console.WriteLine("No trendlines found in the first series of the chart.");
            return;
        }

        // Retrieve the first trendline of the first series
        Trendline trendline = chart.NSeries[0].TrendLines[0];

        // The equation text is not directly exposed via a property.
        // As a common practice, the equation can be stored in the trendline's Name.
        // If the Name property was set to the equation when the trendline was created,
        // we can retrieve it here. Adjust this logic if your workbook stores the equation elsewhere.
        string equationText = trendline.Name ?? string.Empty;

        // Store the retrieved equation text into a cell (e.g., C1)
        worksheet.Cells["C1"].PutValue(equationText);

        // Save the workbook with the updated cell
        string outputPath = "output.xls";
        workbook.Save(outputPath);

        Console.WriteLine($"Trendline equation saved to {outputPath} in cell C1.");
    }
}