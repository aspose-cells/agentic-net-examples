// Title: Handle Axis.GetAxisTexts Exception When Chart.Calculate Is Missing – Aspose.Cells for .NET
// Description: A C# example that creates a workbook, adds a column chart, deliberately calls chart.ValueAxis.GetAxisTexts() before chart.Calculate(), catches the expected exception, then runs Calculate() and successfully retrieves the axis labels. The demo shows proper try‑catch handling and safe chart‑axis text extraction before saving the file.
// Keywords: Aspose.Cells | .NET | C# | Axis.GetAxisTexts | Chart.Calculate | exception handling | chart axis labels | retrieve axis texts | Aspose.Cells chart example | GetAxisTexts before Calculate
// Common Searches: Axis.GetAxisTexts throws exception Aspose.Cells | why need Chart.Calculate before GetAxisTexts | Aspose.Cells chart axis error handling C# | how to get axis labels after chart calculation | C# example GetAxisTexts without Calculate
// Developer Intent: Show developers how to ensure a chart is calculated before accessing its axis texts and how to gracefully handle the exception when this prerequisite is missed.
// Use Cases: Prevent runtime errors in automated report generators that extract chart axis labels. | Provide clear user feedback when GetAxisTexts is called prematurely. | Integrate safe axis‑label retrieval into existing Aspose.Cells workflows.
// AI Prompts: Write C# code that checks if a chart has been calculated; if not, call Chart.Calculate and then retrieve Axis.GetAxisTexts with proper exception handling. | Create a reusable Aspose.Cells method that returns axis labels or a friendly error message when the chart is not yet calculated. | Generate an example demonstrating the exception thrown by Axis.GetAxisTexts before Chart.Calculate and the correct sequence to obtain labels after calculation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // A C# example that creates a workbook, adds a column chart, deliberately calls chart.ValueAxis.GetAxisTexts() before chart.Calculate(), catches the expected exception, then runs Calculate() and successfully retrieves the axis labels. The demo shows proper try‑catch handling and safe chart‑axis text extraction before saving the file.
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

            // -----------------------------------------------------------------
            // Attempt to get axis labels BEFORE calling Chart.Calculate()
            // This should raise an exception because the axis texts are not yet generated.
            // -----------------------------------------------------------------
            try
            {
                string[] labelsBeforeCalc = chart.ValueAxis.GetAxisTexts();
                Console.WriteLine("Axis labels retrieved without calculation (unexpected):");
                foreach (string label in labelsBeforeCalc)
                {
                    Console.WriteLine(label);
                }
            }
            catch (Exception ex)
            {
                // Expected path: inform the user that Calculate() must be called first
                Console.WriteLine("Error retrieving axis texts before calculation: " + ex.Message);
            }

            // Now calculate the chart to generate axis labels
            chart.Calculate();

            // Retrieve axis labels after calculation – this should succeed
            try
            {
                string[] labelsAfterCalc = chart.ValueAxis.GetAxisTexts();
                Console.WriteLine("Axis labels after Chart.Calculate():");
                foreach (string label in labelsAfterCalc)
                {
                    Console.WriteLine(label);
                }
            }
            catch (Exception ex)
            {
                // Any unexpected error will be reported here
                Console.WriteLine("Unexpected error after calculation: " + ex.Message);
            }

            // Save the workbook
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
