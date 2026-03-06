using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCultureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(1500);
            dataSheet.Cells["A3"].PutValue("Orange");
            dataSheet.Cells["B3"].PutValue(2500);
            dataSheet.Cells["A4"].PutValue("Banana");
            dataSheet.Cells["B4"].PutValue(1800);

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create a pivot table based on the data range
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales as data field

            // Set the culture info for the workbook (affects the pivot table)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR"); // French culture

            // Refresh and calculate the pivot table to apply the culture settings
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTableWithCulture.xlsx", SaveFormat.Xlsx);
        }
    }
}