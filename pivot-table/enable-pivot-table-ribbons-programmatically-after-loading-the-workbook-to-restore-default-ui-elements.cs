// Title: Programmatically restore the default PivotTable ribbon UI in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that sets Workbook.RibbonXml to null, turns on PivotTable.EnableWizard and EnableFieldList for every pivot table, and enables RefreshDataOnOpeningFile. | Provide an example that iterates through all worksheets, refreshes each PivotTable after activating the default UI elements, and saves the modified workbook.
// Common Searches: Aspose.Cells how to enable pivot wizard and field list for all pivot tables | C# code to remove custom ribbon XML from Excel workbook with Aspose.Cells | Make pivot field list visible by default in an existing .xlsx using Aspose.Cells | Automatically refresh pivot tables on opening file with Aspose.Cells .NET
// Tags: reset RibbonXml Aspose.Cells .NET | activate pivot wizard field list C# | show pivot field list workbook | configure pivot refresh on file open | default pivot UI settings Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDemo
{
    // Loads an existing workbook, clears any custom Ribbon XML, makes the pivot field list visible, enables the wizard and field list for each PivotTable, sets them to refresh when the file opens, refreshes all pivot tables, and saves the updated workbook.
    public class Program
    {
        public static void Main()
        {
            const string inputPath = "InputWithPivot.xlsx";
            const string outputPath = "OutputWithDefaultPivotUI.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook that contains pivot tables
                Workbook workbook = new Workbook(inputPath);

                // Use the default Ribbon UI (remove any custom Ribbon XML)
                workbook.RibbonXml = null;

                // Ensure the pivot field list is visible at the workbook level
                workbook.Settings.HidePivotFieldList = false;

                // Iterate through all worksheets and their pivot tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        // Enable UI elements for the pivot table
                        pt.EnableWizard = true;
                        pt.EnableFieldList = true;

                        // Refresh data automatically when the file is opened
                        pt.RefreshDataOnOpeningFile = true;
                    }

                    // Apply changes to all pivot tables in the current worksheet
                    sheet.RefreshPivotTables();
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
