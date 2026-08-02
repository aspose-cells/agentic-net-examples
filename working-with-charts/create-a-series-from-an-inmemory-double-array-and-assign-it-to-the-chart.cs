using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // In‑memory double array
        double[] data = new double[] { 10.5, 20.75, 30.0, 40.25, 50.5 };

        // Import the array vertically starting at cell A1 (row 0, column 0)
        sheet.Cells.ImportArray(data, 0, 0, true);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Create a series that uses the imported double values as its Y‑values
        // The range refers to the cells where the array was placed (A1:A5)
        chart.NSeries.Add("=Sheet1!$A$1:$A$5", true);

        // Save the workbook
        workbook.Save("ChartFromDoubleArray.xlsx");
    }
}