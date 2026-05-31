using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing workbook that contains the pie chart
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Locate the first chart of type Pie
        Chart pieChart = null;
        foreach (Chart c in sheet.Charts)
        {
            if (c.Type == ChartType.Pie)
            {
                pieChart = c;
                break;
            }
        }

        // If a pie chart is found, change its type to Doughnut
        if (pieChart != null)
        {
            pieChart.Type = ChartType.Doughnut;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}