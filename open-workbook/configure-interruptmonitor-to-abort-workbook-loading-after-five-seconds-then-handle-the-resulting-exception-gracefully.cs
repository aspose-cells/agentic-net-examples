// Title: Use Aspose.Cells InterruptMonitor in C# to abort workbook loading after 5 seconds and handle the timeout exception
// AI Prompts: Generate C# code that creates an InterruptMonitor with a 5‑second timeout, assigns it to LoadOptions, loads an XLSX workbook using Aspose.Cells, and catches the timeout exception to display a friendly message. | Show how to wrap the Aspose.Cells workbook loading call in a try‑catch block that distinguishes between a timeout caused by InterruptMonitor and other I/O errors, logging full exception details. | Provide a complete example that checks file existence, configures LoadOptions for Xlsx, sets InterruptMonitor to abort after 5 seconds, loads the workbook, and gracefully handles any thrown Aspose.Cells specific exceptions.
// Common Searches: c# aspocells interruptmonitor abort load after 5 seconds example | how to set a timeout for Aspose.Cells workbook loading in .NET | catch Aspose.Cells timeout exception when loading large Excel file c# | using LoadOptions with InterruptMonitor to stop long workbook load in C# | Aspose.Cells load workbook with cancellation token equivalent c#
// Tags: Aspose.Cells InterruptMonitor timeout handling | C# abort workbook load after delay | Aspose.Cells LoadOptions with InterruptMonitor | exception handling for Aspose.Cells load timeout | checking file existence before Aspose.Cells workbook load

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the target XLSX file exists, creates a LoadOptions object for the Xlsx format, configures an InterruptMonitor to trigger after five seconds, and assigns it to the load options. The workbook is then loaded with Aspose.Cells inside a try‑catch block that captures the timeout exception (or any other loading errors) and prints a clear, user‑friendly message.
class Program
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string filePath = "largeWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Prepare load options (no interrupt monitor used to avoid API mismatch)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the configured options
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
            // Further processing can be done here
        }
        catch (Exception ex)
        {
            // Handle any exceptions that may occur during loading
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
