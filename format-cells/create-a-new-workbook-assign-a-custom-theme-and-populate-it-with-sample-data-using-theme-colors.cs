// Title: Create a new Excel workbook with a custom theme and style header and data rows using Aspose.Cells for .NET (C#)
// Common Searches: how to define and apply a custom theme with accent colors in Aspose.Cells C# | Aspose.Cells example for styling header row with custom theme colors | populate sample data and apply row background colors using Aspose.Cells .NET | C# code to auto‑fit columns after applying a custom theme in Aspose.Cells | save workbook with custom theme colors using Aspose.Cells for .NET
// Tags: theme color palette definition Aspose.Cells | cell background styling with accent colors C# | header row formatting Excel Aspose.Cells | sample data insertion into workbook C# | auto fit columns after styling Aspose.Cells | export workbook to .xlsx using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, defines six custom accent colors, adds a header row styled with the first accent (teal), inserts three rows of sample items styled with the second accent (orange), auto‑fits the columns, and saves the file as CustomThemeSample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Define custom accent colors --------------------
            Color accent1 = Color.FromArgb(0, 128, 128);   // Teal
            Color accent2 = Color.FromArgb(255, 165, 0);   // Orange
            Color accent3 = Color.FromArgb(128, 0, 128);   // Purple
            Color accent4 = Color.FromArgb(0, 0, 255);     // Blue
            Color accent5 = Color.FromArgb(255, 0, 0);     // Red
            Color accent6 = Color.FromArgb(0, 255, 0);     // Green

            // -------------------- Populate sample data --------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            // Style for header using Accent1 background
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = accent1; // Teal background
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.Color = Color.White;
            headerStyle.Font.IsBold = true;

            StyleFlag flag = new StyleFlag { All = true };
            cells.CreateRange("A1:C1").ApplyStyle(headerStyle, flag);

            // Sample data rows
            string[] items = { "Apple", "Banana", "Cherry" };
            int[] quantities = { 10, 20, 15 };
            double[] prices = { 0.5, 0.3, 0.8 };

            for (int i = 0; i < items.Length; i++)
            {
                int row = i + 1; // Data starts at row 2 (index 1)

                cells[row, 0].PutValue(items[i]);
                cells[row, 1].PutValue(quantities[i]);
                cells[row, 2].PutValue(prices[i]);

                // Row style using Accent2 background
                Style rowStyle = workbook.CreateStyle();
                rowStyle.ForegroundColor = accent2; // Orange background
                rowStyle.Pattern = BackgroundType.Solid;
                rowStyle.Font.Color = Color.Black;

                cells.CreateRange(row, 0, 1, 3).ApplyStyle(rowStyle, flag);
            }

            // Adjust column widths
            sheet.AutoFitColumns();

            // -------------------- Save the workbook --------------------
            string outputPath = "CustomThemeSample.xlsx";

            // Ensure the directory exists (useful if a path is provided)
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
