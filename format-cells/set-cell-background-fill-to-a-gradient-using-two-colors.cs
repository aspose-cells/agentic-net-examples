using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for GradientStyleType enum

namespace AsposeCellsGradientDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a style object
                Style style = workbook.CreateStyle();

                // Set a two‑color gradient on the style (LightSkyBlue → DarkBlue, horizontal)
                style.SetTwoColorGradient(
                    Color.LightSkyBlue,          // first gradient color
                    Color.DarkBlue,              // second gradient color
                    GradientStyleType.Horizontal, // gradient direction
                    1);                          // variant (1‑4)

                // Apply the style to a specific cell
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("Gradient Cell");
                cell.SetStyle(style);

                // Adjust row height and column width for better visibility
                worksheet.Cells.SetRowHeight(0, 40);
                worksheet.Cells.SetColumnWidth(0, 30);

                // Save the workbook
                string outputPath = "CellTwoColorGradient.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}