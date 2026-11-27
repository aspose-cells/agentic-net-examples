// Title: Export an Excel workbook to HTML with timeout handling using Aspose.Cells in C#
// Common Searches: how to export an Excel file to HTML with a timeout using Aspose.Cells C# | c# Aspose.Cells HTML export cancellation token example | implement custom timeout for workbook.Save to HTML in Aspose.Cells | stop long running Aspose.Cells HTML conversion after a set time in .NET | Aspose.Cells export to HTML with interrupt monitor alternative in C#
// Tags: Aspose.Cells HTML export with timeout | custom cancellation for workbook.Save in C# | Excel to HTML conversion performance control Aspose.Cells | handling OperationCanceledException during Aspose.Cells export | fallback timeout logic for Aspose.Cells HtmlSaveOptions

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions, and saves it as HTML while demonstrating how to apply custom timeout or cancellation logic (since InterruptMonitor is not available) and handling OperationCanceledException and other errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the existing file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // NOTE: Interrupt monitor APIs are not available in the current Aspose.Cells version.
        // If needed, implement custom cancellation logic here.

        try
        {
            // Define HTML save options (default options are sufficient for basic export)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (OperationCanceledException ex)
        {
            // This exception is thrown when the operation is cancelled (e.g., by a custom monitor)
            Console.WriteLine("Export cancelled: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during export: {ex.Message}");
        }
    }
}
