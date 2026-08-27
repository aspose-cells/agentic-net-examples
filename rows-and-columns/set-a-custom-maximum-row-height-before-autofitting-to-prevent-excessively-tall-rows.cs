// Title: How to cap row height before auto‑fitting rows with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets a maximum row height using AutoFitterOptions before invoking Worksheet.AutoFitRows in Aspose.Cells. | Show an example that limits auto‑fit row expansion to 50 points while preserving rows that already have a custom height in a .NET Excel workbook.
// Common Searches: C# Aspose.Cells set max row height when auto fitting rows | limit row height after AutoFitRows in Aspose.Cells .NET | AutoFitterOptions OnlyAuto true usage example for Excel row height | prevent rows from becoming excessively tall with Aspose.Cells auto‑fit
// Tags: row height cap with AutoFitterOptions | Worksheet.AutoFitRows limit row height | C# set row height before auto fit | Aspose.Cells prevent tall rows | Excel row height maximum using Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, writes long wrapped text to cell A1, optionally sets an initial row height, configures AutoFitterOptions with MaxRowHeight = 50 points and OnlyAuto = true, auto‑fits all rows respecting the maximum height, and saves the file as MaxRowHeightDemo.xlsx.
    class SetMaxRowHeightBeforeAutoFit
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add long text that would normally cause a very tall row
            worksheet.Cells["A1"].PutValue("This is a very long piece of text that will demonstrate how to limit the row height when auto‑fitting. It contains multiple sentences and should wrap across several lines.");
            // Enable text wrapping so the row height is affected by the content
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Optionally set an initial small row height
            worksheet.Cells.SetRowHeight(0, 10);

            // Create AutoFitterOptions and set the maximum row height (in points)
            AutoFitterOptions options = new AutoFitterOptions
            {
                MaxRowHeight = 50,   // Limit maximum row height to 50 points
                OnlyAuto = true      // Apply only to rows without custom height
            };

            // Auto‑fit all rows using the options (auto‑fit rule)
            worksheet.AutoFitRows(options);

            // Save the workbook (save rule)
            workbook.Save("MaxRowHeightDemo.xlsx");
        }
    }
}
