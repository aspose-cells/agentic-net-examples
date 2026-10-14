// Title: Set worksheet zoom to 150% and freeze the first four rows using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adjusts the worksheet zoom to 150% and freezes rows 1‑4 across the entire used range. | Generate a .NET example that creates a workbook, applies a custom view with a specific zoom level, and uses FreezePanes to lock the top four rows.
// Common Searches: Aspose.Cells C# set zoom level and freeze top rows | How to freeze first four rows after setting zoom in Aspose.Cells .NET | C# Aspose.Cells FreezePanes with custom view zoom | Set worksheet zoom to 150% and freeze panes using Aspose.Cells | Freeze rows while preserving zoom in Excel file with Aspose.Cells
// Tags: adjust zoom level Aspose.Cells C# | freeze top rows Aspose.Cells | custom view Excel Aspose.Cells | FreezePanes used range C# | apply zoom and freeze panes Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, sets the first worksheet's zoom to 150%, freezes the first four rows across the used range with FreezePanes, and saves the file as CustomViewAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the zoom level for the worksheet (e.g., 150%)
            sheet.Zoom = 150;

            // Freeze the first 4 rows (row index is zero‑based, so use 4)
            // Use the overload that specifies total rows and columns to avoid version issues
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(4, 0, totalRows, totalColumns);

            // Define output path
            string outputPath = "CustomViewAndFreeze.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
