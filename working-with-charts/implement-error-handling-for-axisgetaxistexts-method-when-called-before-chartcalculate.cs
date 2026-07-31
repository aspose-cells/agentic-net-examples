// Title: C# – Handle Axis.GetAxisTexts Exception When Chart.Calculate Has Not Been Called (Aspose.Cells)
// Description: Demonstrates how to safely call chart.ValueAxis.GetAxisTexts() in Aspose.Cells for .NET by catching the exception that occurs before Chart.Calculate(), then calculating the chart, retrieving the axis labels, and saving the workbook.
// Keywords: Aspose.Cells | Axis.GetAxisTexts | Chart.Calculate | C# | exception handling | chart axis labels | value axis texts | Aspose.Cells .NET | chart lifecycle | error handling
// Common Searches: Axis.GetAxisTexts throws exception before chart calculation | how to catch GetAxisTexts error Aspose.Cells | retrieve chart axis labels after Calculate in C# | Aspose.Cells chart axis exception handling | GetAxisTexts usage example
// Developer Intent: Show how to wrap Axis.GetAxisTexts in a try‑catch block, call Chart.Calculate, then obtain axis texts without errors.
// Use Cases: Prevent application crashes by handling the pre‑calculation exception for Axis.GetAxisTexts. | Generate accurate value‑axis labels after the chart is calculated for reporting or UI display. | Demonstrate a complete chart workflow: data setup, error‑handled label retrieval, and workbook saving.
// AI Prompts: Write C# code using Aspose.Cells that catches the exception from chart.ValueAxis.GetAxisTexts() when Chart.Calculate() hasn't been executed. | Show the correct sequence to calculate a chart, retrieve value axis texts, and output them without errors. | Provide an example that includes error handling for Axis.GetAxisTexts and saves the workbook after successful label extraction.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to safely call chart.ValueAxis.GetAxisTexts() in Aspose.Cells for .NET by catching the exception that occurs before Chart.Calculate(), then calculating the chart, retrieving the axis labels, and saving the workbook.
    public class AxisGetAxisTextsErrorHandlingDemo
    {
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

            // Attempt to retrieve axis labels before calling Calculate()
            try
            {
                // This call is expected to throw because the chart hasn't been calculated yet
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

            // Properly calculate the chart to generate axis labels
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
                // Any unexpected errors will be caught here
                Console.WriteLine("Error retrieving axis texts after calculation: " + ex.Message);
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
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

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                AxisGetAxisTextsErrorHandlingDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}
