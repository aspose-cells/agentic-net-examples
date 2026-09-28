// Title: Set workbook theme Accent1 color as solid fill for cells imported from a DataTable using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a DataTable into a worksheet and applies the workbook's Accent1 theme color as a solid background to every cell with Aspose.Cells. | Show how to retrieve the Accent1 color from a workbook's theme and use it to style a range of cells after importing data in Aspose.Cells for .NET. | Create a method that formats all cells of an imported DataTable with a solid fill based on the workbook's Accent1 theme color using Aspose.Cells.
// Common Searches: Aspose.Cells C# apply workbook theme Accent1 background to imported DataTable cells | how to use theme colors for cell fill in Aspose.Cells .NET | set solid fill color from workbook theme for a range after data import Aspose.Cells | C# Aspose.Cells change cell background to Accent1 after populating worksheet | retrieve Accent1 color from workbook theme in Aspose.Cells
// Tags: theme accent1 fill Aspose.Cells | solid background color cell Aspose.Cells | format datatable import cells Aspose.Cells | workbook theme color cell style .NET | Aspose.Cells cell style accent color

using System;
using System.Data;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, imports a DataTable into the first worksheet, then iterates over all populated cells and sets a solid fill using the workbook’s Accent1 theme color (or a fallback RGB) before saving the file as ExportedData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Simulate data import (e.g., from a DataTable)
            DataTable table = GetSampleData();

            // Manually import the DataTable into the worksheet starting at cell A1
            int startRow = 0;
            int startCol = 0;

            // Write header row
            for (int c = 0; c < table.Columns.Count; c++)
            {
                sheet.Cells[startRow, startCol + c].PutValue(table.Columns[c].ColumnName);
            }
            startRow++; // Move to first data row

            // Write data rows
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    sheet.Cells[startRow + r, startCol + c].PutValue(table.Rows[r][c]);
                }
            }

            // Define a fallback Accent1 color (replace with actual theme color if available)
            Color accent1Color = Color.FromArgb(0, 112, 192); // Example accent color

            // Apply the Accent1 color as fill to all imported cells
            int totalRows = table.Rows.Count + 1; // +1 for header row
            int totalCols = table.Columns.Count;

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    Style style = cell.GetStyle();

                    // Set solid fill with Accent1 color
                    style.ForegroundColor = accent1Color;
                    style.Pattern = BackgroundType.Solid;

                    cell.SetStyle(style);
                }
            }

            // Save the workbook (lifecycle rule: save)
            string outputPath = "ExportedData.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to generate sample data for import
    static DataTable GetSampleData()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Quantity", typeof(int));

        dt.Rows.Add(1, "Apple", 50);
        dt.Rows.Add(2, "Banana", 30);
        dt.Rows.Add(3, "Cherry", 20);

        return dt;
    }
}
