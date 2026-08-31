// Title: Catch and handle the exception from Axis.GetAxisTexts when called before Chart.Calculate in Aspose.Cells for .NET
// AI Prompts: Write C# code that checks whether a chart has been calculated before invoking chart.ValueAxis.GetAxisTexts, and logs a friendly message if not. | Show how to wrap chart.ValueAxis.GetAxisTexts in a try‑catch block to gracefully handle the exception thrown when Chart.Calculate hasn't been executed. | Generate a complete Aspose.Cells example that demonstrates retrieving axis labels after Chart.Calculate and handling the pre‑calculation error.
// Common Searches: Aspose.Cells Axis.GetAxisTexts throws InvalidOperationException if chart not calculated | how to prevent exception when calling GetAxisTexts on a chart before Calculate in C# | example of error handling for chart axis text retrieval in Aspose.Cells .NET | retrieving value axis labels after chart.Calculate Aspose.Cells | C# Aspose.Cells chart.Calculate required for GetAxisTexts method
// Tags: exception handling for Axis.GetAxisTexts Aspose.Cells | chart.Calculate prerequisite for GetAxisTexts | value axis label retrieval error handling C# | Aspose.Cells column chart axis text extraction | try‑catch around GetAxisTexts Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data and a column chart, then attempts to call chart.ValueAxis.GetAxisTexts before calling chart.Calculate. The expected exception is caught and logged. After invoking chart.Calculate, the axis labels are retrieved successfully and printed. The workbook is saved, and all operations are wrapped in try‑catch blocks to demonstrate robust error handling for Axis.GetAxisTexts.
    public class AxisGetAxisTextsErrorHandlingDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }

        public static void Run()
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

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Attempt to get axis labels before calling Calculate()
            try
            {
                // This call is expected to throw because the chart has not been calculated yet
                string[] preCalcLabels = chart.ValueAxis.GetAxisTexts();
                Console.WriteLine("Axis labels retrieved before calculation (unexpected):");
                foreach (string label in preCalcLabels)
                {
                    Console.WriteLine(label);
                }
            }
            catch (Exception ex)
            {
                // Handle the expected exception gracefully
                Console.WriteLine("Error retrieving axis texts before calculation: " + ex.Message);
            }

            // Now calculate the chart to generate axis labels
            chart.Calculate();

            // Retrieve axis labels after calculation
            try
            {
                string[] labels = chart.ValueAxis.GetAxisTexts();
                Console.WriteLine("Value Axis Labels after calculation:");
                foreach (string label in labels)
                {
                    Console.WriteLine(label);
                }
            }
            catch (Exception ex)
            {
                // Any unexpected errors will be reported here
                Console.WriteLine("Error retrieving axis texts after calculation: " + ex.Message);
            }

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            try
            {
                workbook.Save("AxisGetAxisTextsErrorHandlingDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
