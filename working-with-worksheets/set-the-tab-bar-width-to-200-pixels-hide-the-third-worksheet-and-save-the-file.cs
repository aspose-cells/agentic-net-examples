// Title: Hide the third worksheet and save an Excel workbook as XLSX using Aspose.Cells for .NET (TabBarWidth not supported)
// Common Searches: how to hide a specific worksheet in Aspose.Cells for .NET | Aspose.Cells .NET hide third sheet programmatically | save workbook to xlsx after hiding worksheets using Aspose.Cells C# | set tab bar width in Excel with Aspose.Cells .NET | does Aspose.Cells .NET support TabBarWidth property
// Tags: worksheet visibility management Aspose.Cells .NET | export workbook to xlsx using Aspose.Cells C# | TabBarWidth property unsupported Aspose.Cells .NET | Excel UI customization limitations Aspose.Cells | control sheet visibility Aspose.Cells example

using System;
using Aspose.Cells;

// The example creates a new Workbook, notes that the TabBarWidth property is unavailable in the .NET API, hides the third worksheet (index 2) when at least three sheets exist, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // NOTE: The TabBarWidth property is not available in the .NET API.
            // If needed, adjust UI settings via other means or omit this step.

            // Hide the third worksheet (zero‑based index 2) if it exists
            if (workbook.Worksheets.Count > 2)
            {
                workbook.Worksheets[2].IsVisible = false;
            }

            // Save the workbook to a file
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
