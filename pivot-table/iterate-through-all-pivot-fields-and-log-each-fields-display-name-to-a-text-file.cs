using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFieldLogger
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 600;

            sheet.Cells["A5"].Value = "Furniture";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 400;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (row, column, data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that fields are initialized
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Open a text file for logging the display names of all pivot fields
            using (StreamWriter writer = new StreamWriter("PivotFields.txt"))
            {
                // Iterate through the BaseFields collection which contains all fields of the pivot table
                PivotFieldCollection allFields = pivotTable.BaseFields;

                foreach (PivotField field in allFields)
                {
                    // Write the display name of each field to the file
                    writer.WriteLine(field.DisplayName);
                }
            }

            // Save the workbook (optional, just to keep the workbook file)
            workbook.Save("PivotTableWithLoggedFields.xlsx");
        }
    }
}