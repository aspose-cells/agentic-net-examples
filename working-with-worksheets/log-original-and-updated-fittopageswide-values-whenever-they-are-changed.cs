// Title: Log FitToPagesWide Changes in Aspose.Cells for .NET
// Description: Shows how to read, modify, and log the FitToPagesWide property of a worksheet's PageSetup with Aspose.Cells for .NET, printing values before and after each change and then saving the workbook.
// Keywords: Aspose.Cells | FitToPagesWide | PageSetup | C# | .NET | worksheet pagination | property logging | console output | workbook save
// Common Searches: Aspose.Cells log FitToPagesWide value | C# record page setup changes Aspose.Cells | how to output original and new FitToPagesWide | track pagination settings in Aspose.Cells workbook | save workbook after changing FitToPagesWide
// Developer Intent: The developer wants to display the original FitToPagesWide setting and each subsequent value whenever the property is updated.
// Use Cases: Debug page‑layout scaling by printing FitToPagesWide before and after adjustments. | Create an audit trail of pagination settings for compliance reporting. | Verify that dynamic page‑width scaling is applied correctly prior to saving the file.
// AI Prompts: Generate a reusable C# method that logs the previous and new FitToPagesWide values whenever the property is set on a Worksheet. | Provide code to capture FitToPagesWide changes for all worksheets in a workbook and export the change log to a CSV file. | Explain how to integrate FitToPagesWide logging into a larger Aspose.Cells processing pipeline that also modifies other PageSetup options.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to read, modify, and log the FitToPagesWide property of a worksheet's PageSetup with Aspose.Cells for .NET, printing values before and after each change and then saving the workbook.
    public class FitToPagesWideLoggingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the PageSetup object
                PageSetup pageSetup = worksheet.PageSetup;

                // Log the original FitToPagesWide value (default is 1)
                Console.WriteLine($"Original FitToPagesWide: {pageSetup.FitToPagesWide}");

                // Change the FitToPagesWide value
                pageSetup.FitToPagesWide = 2;

                // Log the updated value
                Console.WriteLine($"Updated FitToPagesWide: {pageSetup.FitToPagesWide}");

                // Change it again to demonstrate multiple logs
                int previousValue = pageSetup.FitToPagesWide;
                pageSetup.FitToPagesWide = 3;
                Console.WriteLine($"FitToPagesWide changed from {previousValue} to {pageSetup.FitToPagesWide}");

                // Save the workbook (lifecycle: save)
                workbook.Save("FitToPagesWideLoggingDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
