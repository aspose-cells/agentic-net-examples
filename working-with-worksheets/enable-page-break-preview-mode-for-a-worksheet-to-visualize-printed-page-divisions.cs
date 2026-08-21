// Title: Enable Page Break Preview Mode for a Worksheet Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to turn on page‑break preview for the first worksheet, set a 100 % zoom level, verify the settings, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# page break preview | worksheet IsPageBreakPreview true | set worksheet zoom Aspose.Cells | export workbook with page break view | Aspose.Cells .NET example | Excel page break visualization | C# Aspose.Cells workbook save
// Common Searches: how to enable page break preview in Aspose.Cells .NET | C# set worksheet zoom and page break view | Aspose.Cells enable IsPageBreakPreview | preview printed pages with Aspose.Cells | Aspose.Cells example for page break preview
// Developer Intent: Activate page‑break preview on a worksheet, apply a consistent zoom, and save the file so Excel opens in preview mode.
// Use Cases: Validate pagination of generated reports before distribution. | Create workbooks that open directly in Excel’s Page Break Preview for end‑users. | Automate visual checks of page layout during batch workbook generation.
// AI Prompts: Show C# code to enable page break preview on a specific worksheet with Aspose.Cells. | Provide an Aspose.Cells snippet that sets IsPageBreakPreview to true and adjusts the zoom to 100 % before saving. | Explain how to confirm that page break preview is active and retrieve the current zoom level using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to turn on page‑break preview for the first worksheet, set a 100 % zoom level, verify the settings, and save the workbook with Aspose.Cells for .NET.
    public class EnablePageBreakPreviewDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable page break preview mode
                worksheet.IsPageBreakPreview = true;

                // Set zoom to 100% for better visibility
                worksheet.Zoom = 100;

                // Output the current settings
                Console.WriteLine("IsPageBreakPreview: " + worksheet.IsPageBreakPreview);
                Console.WriteLine("Zoom: " + worksheet.Zoom);

                // Save the workbook
                string outputPath = "EnablePageBreakPreviewDemo_output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnablePageBreakPreviewDemo.Run();
        }
    }
}
