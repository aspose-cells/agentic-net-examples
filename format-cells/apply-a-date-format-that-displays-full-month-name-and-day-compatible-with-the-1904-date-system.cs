// Title: Format a cell to show full month name and day with a custom number format in Aspose.Cells for .NET using the 1904 date system
// Common Searches: Aspose.Cells C# set custom date format mmmm d for cells with 1904 date system | How to apply full month name and day format to an Excel cell using Aspose.Cells .NET | Enable 1904 date system and format date cells in Aspose.Cells C# example | Create and apply a style with a custom number format for dates in Aspose.Cells | Save an Excel workbook with month name day formatting using Aspose.Cells C#
// Tags: custom number format for dates Aspose.Cells | 1904 date system configuration Aspose.Cells | apply cell style C# Aspose.Cells | format Excel date as month name day | save workbook as .xlsx Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, optionally enables the 1904 date system, inserts a DateTime value into cell A1, defines a style with the custom number format "mmmm d" to display the full month name and day, applies the style to the cell, and saves the workbook as DateFormat1904.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable the 1904 date system if the property exists in your Aspose.Cells version.
            // Uncomment the line below when supported:
            // workbook.Settings.Is1904DateSystem = true;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a date value (e.g., July 15, 2023) into cell A1
            sheet.Cells["A1"].PutValue(new DateTime(2023, 7, 15));

            // Create a style with a custom number format that shows full month name and day
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mmmm d";

            // Apply the style to the cell
            sheet.Cells["A1"].SetStyle(dateStyle);

            // Save the workbook
            workbook.Save("DateFormat1904.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
