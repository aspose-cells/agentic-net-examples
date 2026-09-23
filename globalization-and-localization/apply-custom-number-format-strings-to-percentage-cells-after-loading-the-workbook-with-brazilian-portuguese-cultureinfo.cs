// Title: Apply Brazilian Portuguese custom percentage format (0,00%) to all percent cells using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook with Aspose.Cells, set workbook.Settings.CultureInfo to "pt-BR", and replace every cell where style.IsPercent is true with the custom format string "0,00%". | Create a reusable method that takes a Workbook and a format string, then applies that format to all percentage‑styled cells while keeping the pt‑BR culture settings. | Write a snippet that iterates only the used range of each worksheet, updates percentage cells to a locale‑aware format, and saves the modified workbook to a new file.
// Common Searches: how to set Brazilian Portuguese culture for an Aspose.Cells workbook in C# | change percentage number format to use comma decimal separator with Aspose.Cells | apply custom number format to all percent cells in an Excel file using Aspose.Cells .NET | iterate through used range and modify style.IsPercent cells in Aspose.Cells | save workbook after applying locale‑specific percentage format Aspose.Cells
// Tags: set workbook culture pt-BR Aspose.Cells | custom percentage format 0,00% Aspose.Cells | iterate cells and modify style.IsPercent C# | locale-aware number formatting Excel Aspose.Cells | apply custom number format to percent cells .NET

using System.Globalization;
using Aspose.Cells;

// The example loads input.xlsx, assigns the Brazilian Portuguese (pt-BR) CultureInfo to the workbook, scans every cell, and for cells flagged as percentages replaces their style with the custom format "0,00%" before saving to output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply Brazilian Portuguese culture (pt-BR) to the workbook settings
        workbook.Settings.CultureInfo = new CultureInfo("pt-BR");

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Loop through each cell in the used range
            foreach (Cell cell in cells)
            {
                // Get the current style of the cell
                Style style = cell.GetStyle();

                // Check if the cell is formatted as a percentage
                if (style.IsPercent)
                {
                    // Apply a custom number format string for percentages
                    // Using comma as decimal separator according to pt-BR culture
                    style.Custom = "0,00%";

                    // Assign the modified style back to the cell
                    cell.SetStyle(style);
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
