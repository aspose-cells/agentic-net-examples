// Title: Disable PivotTable ShowReportFilterPages in Aspose.Cells for .NET (C#) using reflection
// Description: Shows how to load a workbook, locate the first PivotTable, and set its ShowReportFilterPages property to false via reflection, keeping all report filters on a single worksheet and preserving compatibility with older Aspose.Cells releases.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | ShowReportFilterPages | disable report filter pages | reflection | consolidate pivot filters | single worksheet | legacy Aspose.Cells versions | Workbook.Save
// Common Searches: Aspose.Cells set ShowReportFilterPages false | How to turn off pivot report filter pages in C# | Reflection to modify PivotTable property Aspose.Cells | Consolidate pivot filters onto one sheet Aspose.Cells | ShowReportFilterPages property missing Aspose.Cells
// Developer Intent: Turn off the ShowReportFilterPages option so a PivotTable’s report filters appear on the same sheet instead of separate pages.
// Use Cases: Preparing a report for distribution where extra filter worksheets increase file size. | Processing legacy workbooks created with earlier Aspose.Cells releases that lack the ShowReportFilterPages property. | Automating batch conversion of multiple workbooks to a uniform layout before publishing.
// AI Prompts: Generate C# code that iterates through all PivotTables in a workbook and disables ShowReportFilterPages using reflection. | Create a utility method for Aspose.Cells that checks for the ShowReportFilterPages property and safely sets it to false. | Explain how to handle missing PivotTable properties in older Aspose.Cells versions when customizing pivot layouts.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Reflection;

// Shows how to load a workbook, locate the first PivotTable, and set its ShowReportFilterPages property to false via reflection, keeping all report filters on a single worksheet and preserving compatibility with older Aspose.Cells releases.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the pivot table
        Workbook workbook = new Workbook("source.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (worksheet.PivotTables.Count > 0)
        {
            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Attempt to set the ShowReportFilterPages property to false.
            // This property may not exist in older versions, so use reflection as a safe fallback.
            PropertyInfo prop = typeof(PivotTable).GetProperty("ShowReportFilterPages");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(pivotTable, false);
            }
            else
            {
                // If the property is unavailable, simply avoid calling any ShowReportFilterPage methods.
                // The pivot table will remain on a single worksheet.
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
