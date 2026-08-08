// Title: C# – AutoFitRows with AutoFitterOptions.ForRendering for wrapped text in Aspose.Cells
// Description: Demonstrates how to place a long string in a cell, enable text wrapping, set a narrow column width, and call Worksheet.AutoFitRows with AutoFitterOptions.ForRendering = true so the row height expands correctly and prevents clipping during rendering. The sample also shows how to read the adjusted row height and save the workbook.
// Keywords: Aspose.Cells AutoFitRows | AutoFitterOptions ForRendering | wrap text row height .NET | prevent text clipping Aspose | C# spreadsheet rendering | read row height after autofit | adjust column width Aspose.Cells
// Common Searches: AutoFitRows ForRendering example C# | how to avoid wrapped text clipping Aspose.Cells | auto fit row height for wrapped text .NET | retrieve row height after AutoFitRows | Aspose.Cells rendering row height issue
// Developer Intent: Automatically expand row height for cells with wrapped text so the full content is visible when the worksheet is rendered or exported.
// Use Cases: Prepare worksheets with long wrapped text for PDF or image export without truncation. | Calculate exact row heights after autofit to align custom graphics or reports. | Ensure consistent on‑screen and printed layout for spreadsheets that contain multi‑line cells.
// AI Prompts: Generate C# code that uses Aspose.Cells to auto‑fit rows with AutoFitterOptions.ForRendering on a sheet containing wrapped text and then output the new row height. | Explain the effect of setting AutoFitterOptions.ForRendering to true on row‑height calculation and when this setting should be applied in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AutoFitRowsForRenderingDemo
{
    // Demonstrates how to place a long string in a cell, enable text wrapping, set a narrow column width, and call Worksheet.AutoFitRows with AutoFitterOptions.ForRendering = true so the row height expands correctly and prevents clipping during rendering. The sample also shows how to read the adjusted row height and save the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into cell A1
            worksheet.Cells["A1"].PutValue("This is a long text that needs auto-fitting for rendering purposes. It contains enough characters to wrap over multiple lines when the column width is limited.");

            // Enable text wrapping for the cell
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Optionally set a narrow column width to force wrapping
            worksheet.Cells.SetColumnWidth(0, 15); // width in characters

            // Create AutoFitterOptions and set ForRendering to true
            AutoFitterOptions options = new AutoFitterOptions
            {
                ForRendering = true
            };

            // Auto-fit rows with rendering considerations
            worksheet.AutoFitRows(options);

            // Output the resulting row height in points (optional)
            Console.WriteLine("Row height after AutoFitRows (points): " + worksheet.Cells.GetRowHeight(0));

            // Save the workbook
            workbook.Save("AutoFitRowsForRenderingDemo.xlsx");
        }
    }
}
