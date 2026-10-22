// Title: Apply business status‑based theme accent fill colors to Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set cell background color based on status string | How to map business status values to Excel cell fill colors with Aspose.Cells | Apply solid theme accent fill to a specific worksheet cell in .NET | Programmatically change Excel cell style using Aspose.Cells based on business logic | Aspose.Cells conditional cell background color for Open, Closed, In Progress statuses
// Tags: status‑driven cell fill Aspose.Cells | solid background pattern C# Excel | business status color mapping Aspose.Cells | apply style to worksheet cell .NET | theme accent fill Excel Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to assign a solid theme accent fill color to a specific Excel cell by mapping a business status string (e.g., "Open", "Closed", "In Progress") to a System.Drawing.Color and applying the style with the Aspose.Cells API in C#.
public static class BusinessStatusStyler
{
    /// <param name="worksheet">The worksheet containing the target cell.</param>
    /// <param name="cellName">The name of the cell (e.g., "A1") to style.</param>
    /// <param name="status">The business status value (e.g., "Open", "Closed").</param>
    public static void ApplyStatusFill(Worksheet worksheet, string cellName, string status)
    {
        if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
        if (string.IsNullOrEmpty(cellName)) throw new ArgumentException("Cell name cannot be null or empty.", nameof(cellName));

        // Map business status to a concrete color.
        Color fillColor = status switch
        {
            "Open" => Color.LightGreen,
            "In Progress" => Color.Khaki,
            "Closed" => Color.LightCoral,
            "On Hold" => Color.LightGray,
            _ => Color.LightBlue // Default color for unknown statuses.
        };

        try
        {
            // Retrieve the cell and its current style.
            Cell cell = worksheet.Cells[cellName];
            Style style = cell.GetStyle();

            // Set the fill pattern to solid and apply the chosen color.
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = fillColor;

            // Apply the modified style back to the cell.
            cell.SetStyle(style);
        }
        catch (Exception ex)
        {
            // Log or handle the exception as needed.
            Console.Error.WriteLine($"Error applying status fill to cell {cellName}: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed).
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example usage of the BusinessStatusStyler.
            sheet.Cells["A1"].PutValue("Task 1");
            BusinessStatusStyler.ApplyStatusFill(sheet, "A1", "Open");

            sheet.Cells["B2"].PutValue("Task 2");
            BusinessStatusStyler.ApplyStatusFill(sheet, "B2", "Closed");

            // Save the workbook to a file.
            string outputPath = "BusinessStatusDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
