// Title: How to set specific row heights and freeze the first two rows in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that sets row 1 height to 30 points and row 2 height to 45 points, then freezes those rows. | Show how to call the FreezePanes method after adjusting row heights so the top two rows stay static while scrolling.
// Common Searches: Aspose.Cells C# set height of first row and second row then freeze panes | How to freeze top two rows after customizing row heights in a .NET Excel workbook | C# example for applying custom row heights before using FreezePanes with Aspose.Cells
// Tags: Aspose.Cells row dimension API C# | FreezePanes method row argument example | Excel worksheet freeze panes header rows .NET | adjust row size Aspose.Cells | static header rows Excel Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, assigns custom heights to the first two rows, freezes those rows using FreezePanes, adds sample data, and saves the file as RowHeightAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom heights for specific rows (zero‑based indices)
            sheet.Cells.Rows[0].Height = 30; // First row
            sheet.Cells.Rows[1].Height = 45; // Second row

            // Freeze the first two rows (rows 0 and 1)
            // FreezePanes(row, column, totalRows, totalColumns)
            // row = 2 (first unfrozen row), column = 0, totalRows = 2, totalColumns = 0
            sheet.FreezePanes(2, 0, 2, 0);

            // Add some sample data to visualize the effect
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue("SubHeader 1");
            sheet.Cells["B2"].PutValue("SubHeader 2");
            sheet.Cells["A3"].PutValue("Data 1");
            sheet.Cells["B3"].PutValue("Data 2");

            // Save the workbook to a file
            string outputPath = "RowHeightAndFreeze.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
