using System;
using Aspose.Cells;

namespace AutoFitRowsForRenderingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a long text that will be wrapped in the cell
            worksheet.Cells["A1"].PutValue("This is a very long text that should be wrapped and auto‑fitted for rendering purposes. It contains enough characters to require multiple lines when wrapped.");

            // Enable text wrapping for the cell
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Configure AutoFitterOptions to fit for rendering
            AutoFitterOptions options = new AutoFitterOptions
            {
                ForRendering = true   // Prevents clipping when the sheet is rendered (e.g., to PDF or image)
            };

            // Auto‑fit the row containing the wrapped text using the rendering options
            worksheet.AutoFitRows(options);

            // Optional: display the resulting row height in points
            Console.WriteLine("Row height after AutoFitRows (points): " + worksheet.Cells.GetRowHeight(0));

            // Save the workbook to demonstrate the result
            workbook.Save("AutoFitRowsForRenderingDemo.xlsx");
        }
    }
}