// Title: C# – Map business status strings to Excel theme accent fill colors using Aspose.Cells
// AI Prompts: Create a C# method that takes a status string and returns an Aspose.Cells Style with a solid background color chosen from the workbook's theme accent palette. | Demonstrate iterating over a collection of status values, applying the generated style to each cell, and saving the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: how to assign Excel theme accent colors to cells based on status values in C# Aspose.Cells | C# Aspose.Cells example for mapping business status to cell fill colors | set solid background pattern with specific theme accent in Aspose.Cells workbook | apply conditional fill style without using Excel formulas using Aspose.Cells .NET | create Excel file with colored status rows using Aspose.Cells API
// Tags: status to theme accent fill color Aspose.Cells | solid background style workbook C# | apply cell fill based on string value Aspose.Cells | generate Excel file with status color coding .NET | use BackgroundType.Solid ForegroundColor Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example defines a BusinessStatusColorMapper class that provides a method to generate an Aspose.Cells Style with a solid theme accent background based on a business status string, applies the style to cells in a new workbook, populates sample statuses, and saves the result as BusinessStatusColors.xlsx.
public static class BusinessStatusColorMapper
{
    public static Style GetStatusFillStyle(string status, Workbook workbook)
    {
        // Create a new style from the workbook.
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;

        // Map status values to concrete colors.
        Color fillColor;
        switch (status?.Trim().ToLowerInvariant())
        {
            case "open":
                fillColor = Color.LightBlue;   // Represents Accent1.
                break;
            case "closed":
                fillColor = Color.LightGreen;  // Represents Accent2.
                break;
            case "pending":
                fillColor = Color.Orange;      // Represents Accent3.
                break;
            case "in progress":
                fillColor = Color.LightCoral;  // Represents Accent4.
                break;
            case "on hold":
                fillColor = Color.Plum;        // Represents Accent5.
                break;
            default:
                fillColor = Color.LightGray;   // Fallback (Accent6).
                break;
        }

        // Apply the chosen color to the cell background.
        style.ForegroundColor = fillColor;

        return style;
    }

    /// <summary>
    /// Generates a workbook with sample statuses and applies the corresponding fill colors.
    /// </summary>
    public static void ApplyStatusColors()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Sample statuses.
            string[] statuses = { "Open", "Closed", "Pending", "In Progress", "On Hold", "Unknown" };

            // Populate cells and apply corresponding fill styles.
            for (int i = 0; i < statuses.Length; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue(statuses[i]);

                Style statusStyle = GetStatusFillStyle(statuses[i], workbook);
                cell.SetStyle(statusStyle);
            }

            // Save the workbook.
            string outputPath = "BusinessStatusColors.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error applying status colors: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        BusinessStatusColorMapper.ApplyStatusColors();
    }
}
