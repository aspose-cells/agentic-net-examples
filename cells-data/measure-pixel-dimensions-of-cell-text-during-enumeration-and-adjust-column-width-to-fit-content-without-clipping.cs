// Title: How to auto‑fit Excel columns and add pixel padding using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit column width then add extra pixels | measure cell text width in pixels with Aspose.Cells .NET | set Excel column width in pixels after auto‑fit using Aspose.Cells | prevent text clipping by adding column padding in Aspose.Cells C# | retrieve column pixel width and increase it programmatically Aspose.Cells
// Tags: auto-fit column Aspose.Cells | pixel column width Aspose.Cells | column pixel padding Aspose.Cells | cell text pixel measurement Aspose.Cells | Excel column clipping prevention Aspose.Cells

using Aspose.Cells;
using System;

// Creates a workbook, fills cells, auto‑fits each column, reads the pixel width with GetColumnWidthPixel, adds a small padding (e.g., 5 px) using SetColumnWidthPixel, and saves the file, ensuring text is not clipped.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example data – replace with your own loading logic if needed
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["B1"].PutValue("A much longer text that needs more space");
            sheet.Cells["A2"].PutValue("Another text");
            sheet.Cells["B2"].PutValue("Medium length");

            // Determine the number of columns that contain data
            int totalColumns = sheet.Cells.MaxColumn + 1;

            // Auto‑fit each column based on its content
            for (int i = 0; i < totalColumns; i++)
            {
                sheet.AutoFitColumn(i);
            }

            // Add a small padding (e.g., 5 pixels) to each column
            for (int i = 0; i < totalColumns; i++)
            {
                int currentPixelWidth = sheet.Cells.GetColumnWidthPixel(i);
                sheet.Cells.SetColumnWidthPixel(i, currentPixelWidth + 5);
            }

            // Save the workbook with adjusted column widths
            workbook.Save("AdjustedColumns.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
