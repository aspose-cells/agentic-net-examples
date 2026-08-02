// Title: Remove Legend Entry Fill from All Charts in an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: Loads a template workbook, iterates through every worksheet, chart, and series, and sets each series' LegendEntry.IsTextNoFill property to true, eliminating fill from legend entries before saving the file.
// Keywords: Aspose.Cells legend entry no fill | C# remove chart legend fill | Aspose.Cells set IsTextNoFill | iterate charts workbook Aspose | transparent legend text Aspose.Cells | Excel chart legend customization .NET
// Common Searches: Aspose.Cells remove fill from legend entries | C# set chart legend text no fill for all charts | how to make legend background transparent in Excel using Aspose | iterate through workbook charts and clear legend fill | Aspose.Cells legend entry IsTextNoFill example
// Developer Intent: Programmatically clear the fill of legend entries for every series in all charts of a workbook before applying further styling.
// Use Cases: Standardize corporate Excel templates so chart legends have no background color. | Generate automated reports that require transparent legend entries to match brand guidelines. | Pre‑process existing workbooks to ensure consistent legend appearance before applying a new theme.
// AI Prompts: Generate C# code with Aspose.Cells that sets LegendEntry.IsTextNoFill = true for all series in every chart of a workbook and saves the result. | Show how to loop through worksheets and charts to remove legend entry fill, then change the legend font to bold and blue. | Explain how to extend the sample to also hide the legend border while keeping the fill disabled.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a template workbook, iterates through every worksheet, chart, and series, and sets each series' LegendEntry.IsTextNoFill property to true, eliminating fill from legend entries before saving the file.
class SetLegendEntryNoFill
{
    static void Main()
    {
        // Load the template workbook
        Workbook workbook = new Workbook("Template.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Iterate through all series of the chart
                foreach (Series series in chart.NSeries)
                {
                    // Set the legend entry text to have no fill
                    series.LegendEntry.IsTextNoFill = true;
                }
            }
        }

        // Save the modified workbook
        workbook.Save("Result.xlsx");
    }
}
