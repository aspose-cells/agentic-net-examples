// Title: Safely retrieve chart axis labels in Aspose.Cells C# after calculating the chart
// Common Searches: Aspose.Cells C# get axis labels after chart.Calculate | How to avoid errors when calling Axis.GetAxisTexts before chart calculation | C# extension method for safe Axis.GetAxisTexts usage in Aspose.Cells | Example of null‑check and try/catch around Aspose.Cells chart processing | Validate chart is calculated before extracting CategoryAxis texts in C#
// Tags: chart.Calculate required for Axis.GetAxisTexts | C# safe axis label extraction Aspose.Cells | extension method GetAxisTextsSafe Aspose.Cells | exception handling for missing chart calculation | validate chart calculation before axis access

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example adds a GetAxisTextsSafe extension that checks the Axis argument and returns its labels only after the owning chart has been calculated. The program loads a workbook, verifies the file and chart existence, calls chart.Calculate(), then safely reads the CategoryAxis texts via the extension, outputting each label inside a try/catch block for robust error handling.
public static class AxisExtensions
{
    /// <param name="axis">The Axis instance.</param>
    /// <returns>An array of axis label strings.</returns>
    public static string[] GetAxisTextsSafe(this Axis axis)
    {
        if (axis == null)
            throw new ArgumentNullException(nameof(axis));

        // Directly return the axis texts. Ensure the owning chart has been calculated before invoking.
        return axis.GetAxisTexts();
    }
}

class Program
{
    static void Main()
    {
        const string workbookPath = "Sample.xlsx";

        // Verify that the workbook file exists to avoid FileNotFoundException.
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure a chart exists on the worksheet.
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart.
            Chart chart = sheet.Charts[0];

            // Calculate the chart before accessing axis texts.
            chart.Calculate();

            // Retrieve X axis texts safely.
            Axis xAxis = chart.CategoryAxis;
            string[] xAxisTexts = xAxis.GetAxisTextsSafe();

            // Output the axis texts.
            foreach (string txt in xAxisTexts)
            {
                Console.WriteLine(txt);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
