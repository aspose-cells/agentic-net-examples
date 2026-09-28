// Title: Highlight duplicate values in column B (rows 2‑101) using Aspose.Cells conditional formatting in C#
// AI Prompts: Write C# code that creates a workbook, defines a CellArea for column B rows 2‑101, adds a DuplicateValues FormatCondition with a solid yellow fill, and saves the file using Aspose.Cells. | Generate a reusable method that applies a duplicate‑value conditional formatting rule to any column range and lets the fill color be passed as a parameter, leveraging the Aspose.Cells API. | Show how to change the background color of an existing duplicate‑value conditional formatting rule from yellow to another color in an Aspose.Cells worksheet.
// Common Searches: aspnet c# how to use Aspose.Cells to highlight duplicate entries in a specific column range | example of DuplicateValues conditional formatting for column B rows 2 to 101 with Aspose.Cells | set background color for duplicate cells in Excel using Aspose.Cells .NET library | apply conditional formatting to a single column in an Aspose.Cells workbook programmatically
// Tags: Aspose.Cells duplicate entry rule | C# conditional formatting column range | Excel duplicate cell highlight Aspose.Cells | Yellow fill style for duplicate cells | CellArea definition for conditional formatting

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, defines a CellArea covering column B rows 2‑101, adds a DuplicateValues conditional formatting rule with a solid yellow background, and saves the workbook as DuplicateValuesHighlight.xlsx.
class DuplicateValuesConditionalFormatting
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook with one worksheet

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to check for duplicate values (column B, rows 2‑101)
            int startRow = 1;      // first data row (0‑based)
            int endRow = 100;      // last data row
            int columnIndex = 1;   // column B (0‑based)

            // Define the cell area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = startRow,
                EndRow = endRow,
                StartColumn = columnIndex,
                EndColumn = columnIndex
            };

            // Add a new ConditionalFormatting entry to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var cf = worksheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with the ConditionalFormatting object
            cf.AddArea(area);

            // Add a condition of type DuplicateValues
            int conditionIndex = cf.AddCondition(FormatConditionType.DuplicateValues);
            FormatCondition condition = cf[conditionIndex];

            // Define the style to apply to duplicate cells (yellow background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            condition.Style = style;

            // Define output file path
            string outputPath = "DuplicateValuesHighlight.xlsx";

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
