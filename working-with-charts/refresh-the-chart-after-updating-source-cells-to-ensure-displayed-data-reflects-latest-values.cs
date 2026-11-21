// Title: Refresh an Excel chart after updating source cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# refresh chart after modifying cell values | update Excel chart data programmatically without reopening workbook using Aspose.Cells | how to ensure chart reflects new cell values when saving workbook in Aspose.Cells | C# Aspose.Cells automatic chart refresh on workbook save | refresh Excel chart source range after cell update Aspose.Cells .NET
// Tags: Aspose.Cells chart refresh on workbook save | C# update Excel chart source cells | programmatic Excel chart data synchronization Aspose.Cells | save workbook to apply chart changes Aspose.Cells | automatic chart data refresh Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing workbook, updates cells B2 and B3, and saves the file. Aspose.Cells automatically refreshes any charts on save, ensuring the displayed data reflects the new cell values without additional API calls.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet containing the source data
            Worksheet sheet = workbook.Worksheets[0];

            // Update source cells with new values
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(200);

            // Charts in Aspose.Cells automatically reflect updated data when the workbook is saved.
            // The loop below is kept for possible future chart-specific handling.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    // Placeholder for any chart-specific operations.
                    // No explicit RefreshData call is required for current Aspose.Cells versions.
                }
            }

            // Save the workbook with the updated data and refreshed charts (lifecycle: save)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
