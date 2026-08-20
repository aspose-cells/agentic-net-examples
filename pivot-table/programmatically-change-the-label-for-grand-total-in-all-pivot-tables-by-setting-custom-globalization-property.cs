// Title: C# – Change Grand Total Caption for All PivotTables via Globalization Settings in Aspose.Cells
// Description: Shows how to replace the default "Grand Total" text with a custom label for every PivotTable in a workbook by using SettablePivotGlobalizationSettings, assigning it to Workbook.Settings.GlobalizationSettings, refreshing each pivot, and saving the result.
// Keywords: Aspose.Cells | C# | PivotTable | Grand Total label | globalization settings | SettablePivotGlobalizationSettings | custom caption | pivot table localization | RefreshData | CalculateData
// Common Searches: Aspose.Cells change Grand Total text | globalization settings pivot table Aspose | set custom Grand Total caption .NET | apply same Grand Total label to multiple pivot tables | localize pivot table totals Aspose.Cells
// Developer Intent: Apply a single custom Grand Total caption to all PivotTables in a workbook using Aspose.Cells globalization configuration.
// Use Cases: Generate reports where every pivot table shares an identical Grand Total wording. | Localize the Grand Total term for multilingual Excel exports. | Update the Grand Total label after source data changes without recreating pivots. | Create template workbooks with a predefined Grand Total caption.
// AI Prompts: Write C# code that sets the Grand Total caption to a user‑provided string for every PivotTable in an existing workbook using Aspose.Cells. | Explain the steps to configure SettablePivotGlobalizationSettings and why RefreshData and CalculateData are required. | Show how to verify that the custom Grand Total label appears in each pivot table after the workbook is saved.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Shows how to replace the default "Grand Total" text with a custom label for every PivotTable in a workbook by using SettablePivotGlobalizationSettings, assigning it to Workbook.Settings.GlobalizationSettings, refreshing each pivot, and saving the result.
    public class ChangeGrandTotalLabelForAllPivotTables
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be used by multiple pivot tables
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Electronics");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Electronics");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Furniture");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(800);
            sheet.Cells["A5"].PutValue("Furniture");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(950);

            // Create first pivot table
            int ptIndex1 = sheet.PivotTables.Add("A1:C5", "E2", "PivotTable1");
            PivotTable pt1 = sheet.PivotTables[ptIndex1];
            pt1.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pt1.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pt1.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Create second pivot table on the same data range but different location
            int ptIndex2 = sheet.PivotTables.Add("A1:C5", "E15", "PivotTable2");
            PivotTable pt2 = sheet.PivotTables[ptIndex2];
            pt2.AddFieldToArea(PivotFieldType.Row, 1);      // Region
            pt2.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // -----------------------------------------------------------------
            // Configure custom globalization settings to change the "Grand Total"
            // label for all pivot tables in the workbook.
            // -----------------------------------------------------------------
            // 1. Create an instance of SettablePivotGlobalizationSettings.
            SettablePivotGlobalizationSettings pivotGlobalSettings = new SettablePivotGlobalizationSettings();

            // 2. Set the desired custom text for the Grand Total label.
            pivotGlobalSettings.SetTextOfGrandTotal("My Custom Grand Total");

            // 3. Attach the custom settings to the workbook's globalization settings.
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                PivotSettings = pivotGlobalSettings
            };

            // Refresh and calculate each pivot table so that the new label takes effect.
            foreach (PivotTable pt in sheet.PivotTables)
            {
                try
                {
                    // Refresh the pivot cache data
                    pt.RefreshData();

                    // Recalculate the pivot table
                    pt.CalculateData();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to refresh pivot table '{pt.Name}': {ex.Message}");
                }
            }

            // Save the workbook to verify the result.
            string outputPath = "ChangedGrandTotalLabel.xlsx";

            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'. All pivot tables now display the custom Grand Total label.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
