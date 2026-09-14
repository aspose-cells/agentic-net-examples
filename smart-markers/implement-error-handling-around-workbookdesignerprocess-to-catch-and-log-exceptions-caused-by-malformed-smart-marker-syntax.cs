// Title: Handle and log CellsException when WorkbookDesigner.Process encounters malformed smart markers in Aspose.Cells for .NET
// AI Prompts: Wrap the call to WorkbookDesigner.Process in a try‑catch that catches CellsException, logs ex.Code and ex.Message, then continues execution. | Add a generic catch block after the specific CellsException handler to record unexpected errors without stopping the workbook save. | Enclose the entire Run method in an outer try‑catch to capture fatal errors and output a concise message before exiting.
// Common Searches: Aspose.Cells how to catch CellsException from smart marker processing | log invalid smart marker syntax errors in C# using Aspose.Cells | save workbook even when WorkbookDesigner.Process fails | example of nested try‑catch for smart markers Aspose.Cells .NET
// Tags: smart marker CellsException handling Aspose.Cells | error logging for smart marker processing | detect malformed smart marker syntax | save workbook after processing exception | outer fatal error handling in Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsErrorHandlingDemo
{
    // The example creates a workbook, inserts an intentionally malformed smart marker, sets a dummy data source, and processes the markers inside nested try‑catch blocks. It catches CellsException to log its code and message, captures any other exceptions, and ensures the workbook is saved even if processing fails, with an outer fatal‑error handler for the Run method.
    public class SmartMarkerProcessor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (could also be loaded from a template file)
                Workbook workbook = new Workbook();

                // Add a worksheet and place a deliberately malformed smart marker
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("&=InvalidSmartMarker"); // malformed syntax

                // Initialize the WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Example data source (can be any valid source)
                designer.SetDataSource("Dummy", new string[] { "Value1", "Value2" });

                try
                {
                    // Process the smart markers; this may throw if syntax is invalid
                    designer.Process();
                }
                catch (CellsException ex) // Aspose.Cells specific exception
                {
                    // Log detailed information about the exception
                    Console.WriteLine("A CellsException was caught while processing smart markers.");
                    Console.WriteLine($"Message: {ex.Message}");
                    Console.WriteLine($"Exception Type Code: {ex.Code}");
                }
                catch (Exception ex)
                {
                    // Catch any other unexpected exceptions
                    Console.WriteLine("An unexpected error occurred during processing.");
                    Console.WriteLine($"Message: {ex.Message}");
                }

                // Save the workbook (even if processing failed, the workbook may still be saved)
                workbook.Save("ProcessedOutput.xlsx");
                Console.WriteLine("Workbook saved as ProcessedOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error in SmartMarkerProcessor.Run:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SmartMarkerProcessor.Run();
        }
    }
}
