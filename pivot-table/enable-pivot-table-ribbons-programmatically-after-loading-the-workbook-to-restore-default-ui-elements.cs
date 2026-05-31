using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string inputFile = "InputWithPivot.xlsx";
                const string outputFile = "OutputWithDefaultRibbon.xlsx";

                // Verify that the input workbook exists before attempting to load it
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                    return;
                }

                // Load the workbook that contains pivot tables
                Workbook workbook = new Workbook(inputFile);

                // Restore the default Ribbon UI by clearing any custom Ribbon XML
                workbook.RibbonXml = null;

                // Ensure the pivot field list UI is visible
                workbook.Settings.HidePivotFieldList = false;

                // Enable wizard and field list for each pivot table in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        pt.EnableWizard = true;
                        pt.EnableFieldList = true;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}