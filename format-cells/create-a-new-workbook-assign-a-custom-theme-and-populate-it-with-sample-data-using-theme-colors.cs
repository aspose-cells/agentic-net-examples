// Title: How to create an Excel workbook with a custom theme and apply theme accent colors to headers and rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom theme accent colors for header row | Create Excel file with alternating row colors using custom theme in Aspose.Cells .NET | Apply solid fill style with custom colors to cells in Aspose.Cells workbook | Auto‑fit columns after styling cells with custom theme in Aspose.Cells C# | Save workbook with custom theme colors using Aspose.Cells for .NET
// Tags: custom theme accent colors Aspose.Cells | header style solid fill Aspose.Cells | alternating row colors Aspose.Cells | auto‑fit columns Aspose.Cells | save workbook .xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeExample
{
    // Demonstrates creating a new Workbook, defining custom theme accent colors, styling the header with a solid fill, applying alternating row colors, auto‑fitting columns, and saving the file as CustomThemeSample.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the default worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // -------------------------------------------------
                // Define custom colors (used as theme accents)
                // -------------------------------------------------
                Color accent1 = Color.FromArgb(0x4F, 0x81, 0xBD); // Blue
                Color accent2 = Color.FromArgb(0xC0, 0x50, 0x4D); // Red
                Color accent3 = Color.FromArgb(0x9B, 0xBA, 0x59); // Green
                Color accent4 = Color.FromArgb(0xFF, 0xC0, 0x00); // Orange
                Color accent5 = Color.FromArgb(0x5B, 0x9B, 0xD5); // Light Blue
                Color accent6 = Color.FromArgb(0x70, 0xAD, 0x47); // Dark Green

                // -------------------------------------------------
                // Populate sample data and apply custom colors
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["C1"].PutValue("Quantity");
                sheet.Cells["D1"].PutValue("Price");

                // Create a style for the header
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.Color = Color.White;
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = accent1;          // Background color
                headerStyle.Pattern = BackgroundType.Solid;    // Solid fill

                // Apply the header style to the first row
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[0, col].SetStyle(headerStyle);
                }

                // Sample data rows
                string[,] data = new string[,]
                {
                    { "Apple", "Fruit", "120", "0.50" },
                    { "Banana", "Fruit", "200", "0.30" },
                    { "Carrot", "Vegetable", "150", "0.20" },
                    { "Broccoli", "Vegetable", "80", "0.40" },
                    { "Chicken", "Meat", "60", "5.00" }
                };

                // Populate data and apply alternating row colors
                for (int row = 0; row < data.GetLength(0); row++)
                {
                    for (int col = 0; col < data.GetLength(1); col++)
                    {
                        sheet.Cells[row + 1, col].PutValue(data[row, col]);
                    }

                    // Create a style for the current row
                    Style rowStyle = workbook.CreateStyle();
                    if (row % 2 == 0)
                    {
                        rowStyle.ForegroundColor = accent2;
                    }
                    else
                    {
                        rowStyle.ForegroundColor = accent3;
                    }
                    rowStyle.Pattern = BackgroundType.Solid;

                    // Apply the row style to columns A-D
                    for (int col = 0; col < 4; col++)
                    {
                        sheet.Cells[row + 1, col].SetStyle(rowStyle);
                    }
                }

                // Auto-fit columns for better appearance
                sheet.AutoFitColumns();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "CustomThemeSample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
