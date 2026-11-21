// Title: How to set a uniform column width for all columns in every worksheet using Cells.StandardWidth in Aspose.Cells for .NET
// Common Searches: Aspose.Cells set same column width for all columns in C# | apply uniform column width to every worksheet using Cells.StandardWidth | C# code to set default column width across all sheets in an Excel file with Aspose.Cells | standard width property example for Aspose.Cells workbook columns
// Tags: standard column width setting Aspose.Cells | Cells.StandardWidth usage C# | apply column width to every sheet | automate Excel column width with Aspose | default column width configuration .NET

using Aspose.Cells;

// Loads a workbook, iterates through each worksheet, sets Cells.StandardWidth to 15 characters to give all columns the same width, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        var workbook = new Workbook("input.xlsx");

        // Apply a uniform column width to all columns in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // StandardWidth is measured in characters of the default font.
            // Set to the desired width (e.g., 15 characters).
            sheet.Cells.StandardWidth = 15;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
