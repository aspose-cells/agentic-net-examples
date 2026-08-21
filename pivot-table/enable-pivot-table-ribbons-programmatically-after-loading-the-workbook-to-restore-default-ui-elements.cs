// Title: Enable PivotTable Ribbon, Wizard, and Field List programmatically with Aspose.Cells for .NET (C#)
// Description: Loads an existing XLSX workbook, removes any custom Ribbon XML to restore Excel's default ribbon, makes the PivotFieldList pane visible, and iterates through all worksheets to set each PivotTable's EnableWizard and EnableFieldList properties to true before saving the file.
// Keywords: Aspose.Cells enable pivot ribbon | restore default Excel ribbon Aspose | show pivot field list C# | pivot table wizard enable Aspose.Cells | clear RibbonXml Aspose | pivot table UI programmatically | Aspose.Cells C# pivot settings
// Common Searches: how to reset ribbon UI for pivot tables using Aspose.Cells | enable pivot table wizard and field list in all sheets Aspose.Cells .NET | remove custom RibbonXml and show pivot field list Aspose.Cells | programmatically show pivot field list in Excel workbook C# | restore default ribbon after loading workbook Aspose
// Developer Intent: Reset any custom ribbon configuration and activate the PivotTable wizard and field‑list UI for every pivot table in a loaded workbook.
// Use Cases: A reporting tool loads a workbook saved with a custom ribbon, clears RibbonXml, and returns the standard Excel ribbon for end users. | Generating dynamic dashboards where all pivot tables must expose the wizard and field‑list dialogs without manual user interaction. | Automating the preparation of shared workbooks so recipients can immediately access PivotTable UI features.
// AI Prompts: Write C# code using Aspose.Cells to clear custom RibbonXml, enable the PivotTable wizard and field list for all pivot tables, and save the workbook. | Show an example that loads an .xlsx file, makes the PivotFieldList pane visible, restores the default ribbon UI, and saves the result. | Explain how to iterate over worksheets and pivot tables to set EnableWizard and EnableFieldList properties with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDemo
{
    // Loads an existing XLSX workbook, removes any custom Ribbon XML to restore Excel's default ribbon, makes the PivotFieldList pane visible, and iterates through all worksheets to set each PivotTable's EnableWizard and EnableFieldList properties to true before saving the file.
    public class Program
    {
        public static void Main()
        {
            const string inputPath = "InputWithPivot.xlsx";
            const string outputPath = "OutputWithDefaultRibbon.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook that contains pivot tables
                Workbook workbook = new Workbook(inputPath);

                // Restore the default Ribbon UI by clearing any custom Ribbon XML
                workbook.RibbonXml = null;

                // Ensure the pivot field list UI is visible
                workbook.Settings.HidePivotFieldList = false;

                // Enable UI features for each pivot table in every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        pt.EnableWizard = true;
                        pt.EnableFieldList = true;
                    }
                }

                // Save the modified workbook (default format is XLSX)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
