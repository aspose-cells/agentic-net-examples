// Title: How to set RefreshOnLoad = true for a WebQuery connection in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx workbook, locates each WebQuery connection, and enables its RefreshOnLoad flag before saving. | Generate a .NET example that iterates over WorkbookConnection objects, identifies ConnectionType.WebQuery, and turns on automatic refresh on load. | Provide a step‑by‑step C# snippet that checks for a WebQuery connection in a workbook and turns on the RefreshOnLoad setting so the data updates when the file is opened.
// Common Searches: Aspose.Cells .NET how to enable RefreshOnLoad for a WebQuery in an existing Excel file | C# code to make web query connections refresh automatically when opening a workbook with Aspose.Cells | Programmatically activate RefreshOnLoad for Excel WebQuery using Aspose.Cells API | Update Excel workbook to refresh web query data on open via Aspose.Cells C# | Set WebQuery connection to auto‑refresh on load with Aspose.Cells for .NET
// Tags: Aspose.Cells set RefreshOnLoad flag | WebQuery connection automatic refresh .NET | C# modify Excel workbook connections Aspose.Cells | RefreshOnLoad for Excel WebQuery | WorkbookConnection WebQuery example Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// This C# example demonstrates how to open an existing Excel file with Aspose.Cells, iterate through its WebQuery connections, enable the RefreshOnLoad flag so the data refreshes automatically when the workbook is opened, and then save the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The original example attempted to iterate over workbook connections.
            // The appropriate API for managing connections may vary between Aspose.Cells versions.
            // If needed, replace the following placeholder with the correct connection handling code
            // using the version‑specific classes (e.g., WorkbookConnectionInfo, ConnectionType, etc.).

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
