using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Proceed only if the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart (assumed to be a PivotChart)
            Chart pivotChart = worksheet.Charts[0];

            // Hide the pivot chart legend (pivot controls) via PivotOptions
            pivotChart.PivotOptions.DropZonesVisible = false;
        }

        // Save the workbook with the changes applied
        workbook.Save("Output.xlsx");
    }
}