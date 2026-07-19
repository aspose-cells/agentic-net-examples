// Title: Create a Chart Series from an In‑Memory Double[] with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to import a double[] into worksheet cells, optionally add string labels, create a column chart, bind the imported range as the series Y‑values, set CategoryData for the X‑axis, and save the workbook as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart series | double array import | ImportArray method | NSeries.Add | CategoryData | column chart | Excel generation .NET | in‑memory data charting | Aspose.Cells tutorial
// Common Searches: Aspose.Cells add series from double array | Import double[] to worksheet for chart C# | Set X axis labels for Aspose.Cells chart | Create column chart programmatically Aspose.Cells | How to use NSeries.Add with cell range
// Developer Intent: Generate a chart whose data comes directly from an in‑memory double[] and optional string[] without external files.
// Use Cases: Quickly visualize numeric results stored in a double[] within an Excel workbook. | Provide custom category labels from a string[] and bind them to the chart’s X‑axis. | Automate report creation where charts are built from runtime data structures.
// AI Prompts: Show me how to create a line chart from a double[] and string[] using Aspose.Cells for .NET. | Explain how to refresh an existing chart series with new in‑memory data in Aspose.Cells. | Provide code to add multiple series from different double arrays to the same chart in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to import a double[] into worksheet cells, optionally add string labels, create a column chart, bind the imported range as the series Y‑values, set CategoryData for the X‑axis, and save the workbook as an Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // In‑memory double array that will become the Y values of the series
        double[] yValues = new double[] { 10.5, 20.75, 30.0, 40.25, 50.5 };

        // Import the double array vertically starting at cell A1 (row 0, column 0)
        // Rule: Cells.ImportArray(double[], int, int, bool)
        sheet.Cells.ImportArray(yValues, 0, 0, true);

        // Optional: create category (X) labels as a string array
        string[] xLabels = new string[] { "A", "B", "C", "D", "E" };
        // Import the string array vertically starting at cell B1 (row 0, column 1)
        // Rule: Cells.ImportArray(string[], int, int, bool)
        sheet.Cells.ImportArray(xLabels, 0, 1, true);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series that uses the imported double values (A1:A5)
        // Rule: SeriesCollection.Add(string, bool)
        chart.NSeries.Add("=Sheet1!$A$1:$A$5", true);

        // Assign the category (X) data to the series (B1:B5)
        chart.NSeries.CategoryData = "=Sheet1!$B$1:$B$5";

        // Save the workbook
        workbook.Save("SeriesFromArray.xlsx");
    }
}
