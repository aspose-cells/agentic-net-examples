// Title: Rename Grand Total label for every PivotTable in an Aspose.Cells workbook using SettablePivotGlobalizationSettings (C#)
// AI Prompts: Apply SettablePivotGlobalizationSettings to replace the default Grand Total caption across all pivot tables in a workbook. | Refresh and recalculate each PivotTable after updating the global pivot globalization settings to reflect the new label. | Save the workbook to an XLSX file after customizing the Grand Total text to verify the change.
// Common Searches: how to change the grand total text for all pivot tables in Aspose.Cells C# | Aspose.Cells SettablePivotGlobalizationSettings example for custom grand total label | programmatically refresh pivot tables after modifying globalization settings Aspose.Cells | global pivot settings to rename grand total caption in .xlsx using Aspose.Cells
// Tags: settextofgrandtotal Aspose.Cells | global pivot globalization settings C# | rename grand total label pivot table | refresh all pivot tables Aspose.Cells | customize pivot table captions .xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // The example creates a workbook with sample data, adds two pivot tables, defines a SettablePivotGlobalizationSettings object to replace the default "Grand Total" caption with a custom text, assigns it to the workbook's globalization settings, refreshes and recalculates each pivot table, and saves the result as PivotTables_With_CustomGrandTotal.xlsx.
    public class ChangeGrandTotalLabelForAllPivotTables
    {
        public static void Main()
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
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot tables
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Region");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Electronics");
            dataSheet.Cells["B2"].PutValue("North");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Electronics");
            dataSheet.Cells["B3"].PutValue("South");
            dataSheet.Cells["C3"].PutValue(1500);

            dataSheet.Cells["A4"].PutValue("Furniture");
            dataSheet.Cells["B4"].PutValue("North");
            dataSheet.Cells["C4"].PutValue(800);

            dataSheet.Cells["A5"].PutValue("Furniture");
            dataSheet.Cells["B5"].PutValue("South");
            dataSheet.Cells["C5"].PutValue(950);

            // Create first pivot table
            Worksheet pivotSheet1 = workbook.Worksheets.Add("Pivot1");
            int pivotIndex1 = pivotSheet1.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable1 = pivotSheet1.PivotTables[pivotIndex1];
            pivotTable1.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable1.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivotTable1.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Create second pivot table on the same data sheet
            Worksheet pivotSheet2 = workbook.Worksheets.Add("Pivot2");
            int pivotIndex2 = pivotSheet2.PivotTables.Add("A1:C5", "E3", "PivotTable2");
            PivotTable pivotTable2 = pivotSheet2.PivotTables[pivotIndex2];
            pivotTable2.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable2.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // -----------------------------------------------------------------
            // Set custom globalization settings for the "Grand Total" label
            // -----------------------------------------------------------------
            SettablePivotGlobalizationSettings customPivotSettings = new SettablePivotGlobalizationSettings();
            customPivotSettings.SetTextOfGrandTotal("My Custom Grand Total");
            workbook.Settings.GlobalizationSettings.PivotSettings = customPivotSettings;

            // Refresh and calculate all pivot tables so that the new label takes effect
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    // Refresh the pivot cache and recalculate data
                    pt.RefreshData();
                    pt.CalculateData();
                }
            }

            // Save the workbook (output file demonstrates the changed Grand Total label)
            workbook.Save("PivotTables_With_CustomGrandTotal.xlsx");
        }
    }
}
