// Title: Aspose.Cells .NET – Handle Axis.GetAxisTexts Exception When Chart.Calculate Has Not Been Called
// Description: C# example that creates a workbook, adds data and a column chart, then demonstrates proper error handling for Axis.GetAxisTexts. The code catches the exception thrown when the method is invoked before Chart.Calculate, calls Calculate to generate axis labels, retrieves them successfully, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | Axis.GetAxisTexts | Chart.Calculate | exception handling | chart axis labels | retrieve axis texts | sample code | GitHub | error handling for charts
// Common Searches: Aspose.Cells Axis.GetAxisTexts throws exception before Calculate | how to catch GetAxisTexts error in Aspose.Cells | chart.Calculate required for axis labels Aspose.Cells | C# example Axis.GetAxisTexts exception handling | retrieve chart axis texts after calculation Aspose.Cells
// Developer Intent: Show how to safely obtain chart axis labels by handling the exception that occurs if GetAxisTexts is called before the chart is calculated.
// Use Cases: Wrap Axis.GetAxisTexts in a try‑catch block, call Chart.Calculate on failure, then retry to get labels. | Log the exception and supply fallback axis values when label extraction is attempted too early. | Validate chart readiness in automated reporting pipelines to prevent runtime crashes.
// AI Prompts: Generate a reusable C# method that returns axis texts, automatically calling Chart.Calculate if GetAxisTexts fails. | Provide a concise snippet showing try‑catch handling for Axis.GetAxisTexts with a default label fallback. | Explain why Axis.GetAxisTexts depends on Chart.Calculate and how to programmatically verify chart readiness before accessing axis data.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds data and a column chart, then demonstrates proper error handling for Axis.GetAxisTexts. The code catches the exception thrown when the method is invoked before Chart.Calculate, calls Calculate to generate axis labels, retrieves them successfully, and saves the workbook.
    public class AxisGetAxisTextsErrorHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(8000);
                worksheet.Cells["B3"].PutValue(4000);
                worksheet.Cells["B4"].PutValue(-8000);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set chart data range
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Attempt to get axis texts before calling Calculate()
                try
                {
                    string[] preCalcLabels = chart.ValueAxis.GetAxisTexts();
                    Console.WriteLine("Axis labels retrieved before Calculate (unexpected):");
                    foreach (string label in preCalcLabels)
                    {
                        Console.WriteLine(label);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error retrieving axis labels before Calculate: " + ex.Message);
                }

                // Calculate the chart to generate axis labels
                chart.Calculate();

                // Retrieve axis texts after calculation
                try
                {
                    string[] postCalcLabels = chart.ValueAxis.GetAxisTexts();
                    Console.WriteLine("Axis labels after Calculate:");
                    foreach (string label in postCalcLabels)
                    {
                        Console.WriteLine(label);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error retrieving axis labels after Calculate: " + ex.Message);
                }

                // Save the workbook
                workbook.Save("AxisGetAxisTextsDemo.xlsx");
                Console.WriteLine("Workbook saved as AxisGetAxisTextsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AxisGetAxisTextsErrorHandlingDemo.Run();
        }
    }
}
