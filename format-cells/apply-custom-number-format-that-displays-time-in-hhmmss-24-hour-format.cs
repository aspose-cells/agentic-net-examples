// Title: C# – Apply Custom Number Format "hh:mm:ss" (24‑hour) with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a DateTime value, define a style with the custom number format "hh:mm:ss", apply the style to a cell, and save the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom time format | C# 24‑hour time format | Aspose.Cells number format hh:mm:ss | .NET Excel custom format | format cells as time Aspose | Aspose.Cells US developers | Aspose.Cells Europe examples
// Common Searches: Aspose.Cells set 24 hour time format C# | how to use custom number format hh:mm:ss in Aspose.Cells | format Excel cell as time using Aspose.Cells .NET | apply custom time style to a cell with Aspose.Cells
// Developer Intent: The developer needs to display a DateTime value in a worksheet cell using the 24‑hour "hh:mm:ss" format.
// Use Cases: Standardizing timestamps in exported schedules or shift rosters. | Generating logs where only the time component must be visible. | Creating dashboards that require consistent 24‑hour time representation across different locales.
// AI Prompts: Show C# code that applies the custom number format "hh:mm:ss" to a range of cells with Aspose.Cells. | Explain how to preserve existing cell styles while adding a 24‑hour time format to selected cells. | Provide a step‑by‑step guide for using Workbook.CreateStyle().Custom to set a time format in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert a DateTime value, define a style with the custom number format "hh:mm:ss", apply the style to a cell, and save the file as an Excel workbook using Aspose.Cells for .NET.
    public class CustomTimeFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a DateTime value that includes time (e.g., 14:30:45)
                worksheet.Cells["A1"].PutValue(new DateTime(2023, 1, 1, 14, 30, 45));

                // Create a style and set the custom number format to 24‑hour time
                Style style = workbook.CreateStyle();
                style.Custom = "hh:mm:ss";

                // Apply the style to the cell
                worksheet.Cells["A1"].SetStyle(style);

                // Save the workbook
                workbook.Save("CustomTimeFormat.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomTimeFormatDemo.Run();
        }
    }
}
