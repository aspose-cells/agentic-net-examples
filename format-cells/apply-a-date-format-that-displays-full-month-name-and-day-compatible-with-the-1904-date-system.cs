// Title: Format a cell to display the full month name and day with a custom "mmmm d" pattern while using the 1904 date system in Aspose.Cells for .NET
// AI Prompts: Write C# code that activates the 1904 date system in an Aspose.Cells workbook and applies a custom date format "mmmm d" to a specific cell. | Show how to create a Style object with the "mmmm d" pattern, assign it to a cell containing a DateTime value, and save the workbook as an .xlsx file. | Demonstrate inserting a DateTime (e.g., January 15, 2023) into a worksheet, enabling the Mac 1904 date system, and formatting the cell to display "January 15".
// Common Searches: Aspose.Cells C# set custom date format "mmmm d" with 1904 date system | How to display month name and day in Excel using Aspose.Cells and Mac 1904 date system | Enable 1904 date system and apply full month name format in Aspose.Cells workbook
// Tags: Aspose.Cells date style mmmm d | 1904 date system activation C# | apply date style to Excel cell | Mac Excel 1904 compatibility Aspose.Cells | save workbook as xlsx with styled date

using System;
using Aspose.Cells;

// Enables the 1904 date system, inserts a DateTime value, creates a style with the custom "mmmm d" format to show the full month name and day, applies the style to the cell, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable the 1904 date system (used by older Mac Excel versions)
        workbook.Settings.Date1904 = true;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a sample date value (e.g., January 15, 2023)
        sheet.Cells["A1"].PutValue(new DateTime(2023, 1, 15));

        // Create a style with a custom date format that shows full month name and day
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "mmmm d";   // e.g., "January 15"

        // Apply the style to the cell containing the date
        sheet.Cells["A1"].SetStyle(dateStyle);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
