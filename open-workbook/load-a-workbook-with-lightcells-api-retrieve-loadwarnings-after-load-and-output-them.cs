// Title: Collect and display load warnings with Aspose.Cells C# using IWarningCallback
// Description: Demonstrates how to implement a custom IWarningCallback, attach it to LoadOptions, load an Excel workbook, and iterate the captured WarningInfo objects to output their type and description.
// Keywords: Aspose.Cells C# load warnings | IWarningCallback example | LoadOptions warning callback | retrieve Excel load warnings | Aspose.Cells workbook loading diagnostics
// Common Searches: Aspose.Cells capture load warnings C# | How to use IWarningCallback with LoadOptions | Get warning types after opening Excel with Aspose.Cells | Collect load warnings in .NET Aspose.Cells | Workbook loading warnings example
// Developer Intent: Capture and log any warnings generated while loading an Excel workbook with Aspose.Cells.
// Use Cases: Log warnings to a file for troubleshooting compatibility issues. | Validate workbook integrity by checking for specific warning types after load. | Show a summary of load warnings in a UI after a user opens a spreadsheet.
// AI Prompts: Provide a C# snippet that uses Aspose.Cells LoadOptions with a custom IWarningCallback to write load warnings to a log file. | Show how to filter collected warnings by WarningType and process only critical ones after loading a workbook. | Explain how to integrate a warning callback into an existing Aspose.Cells loading workflow without impacting performance.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLoadWarningsDemo
{
    // Custom warning callback that stores all warnings in a list
    // Demonstrates how to implement a custom IWarningCallback, attach it to LoadOptions, load an Excel workbook, and iterate the captured WarningInfo objects to output their type and description.
    public class CollectingWarningCallback : IWarningCallback
    {
        // List to hold received warnings
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        // This method is called by Aspose.Cells whenever a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Store the warning for later processing
            Warnings.Add(warningInfo);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Create a custom warning callback instance
            var warningCallback = new CollectingWarningCallback();

            // Initialize LoadOptions and assign the warning callback
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = warningCallback;

            // Load the workbook using the constructor that accepts file path and LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // After loading, output all collected warnings
            Console.WriteLine("Load Warnings:");
            if (warningCallback.Warnings.Count == 0)
            {
                Console.WriteLine("No warnings were generated during load.");
            }
            else
            {
                foreach (var warning in warningCallback.Warnings)
                {
                    Console.WriteLine($"- Type: {warning.WarningType}, Description: {warning.Description}");
                }
            }

            // (Optional) Dispose the workbook if no further processing is needed
            workbook.Dispose();
        }
    }
}
