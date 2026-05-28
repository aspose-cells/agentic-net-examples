using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsDemo
{
    class DisableExpandCollapseDemo
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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
            sheet.Name = "Data";

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "A";
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["A3"].Value = "B";
            sheet.Cells["B3"].Value = 200;
            sheet.Cells["A4"].Value = "A";
            sheet.Cells["B4"].Value = 150;

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Disable expand/collapse (drill) buttons in the UI and when printing
            pivotTable.ShowDrill = false;   // hide expand/collapse buttons on screen
            pivotTable.PrintDrill = false;  // prevent printing of drill indicators

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            string outputPath = "PivotTable_NoExpandCollapse.xlsx";
            workbook.Save(outputPath);
        }
    }
}