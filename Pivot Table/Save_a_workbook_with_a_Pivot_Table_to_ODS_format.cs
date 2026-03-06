using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue(3000);

        // Create a pivot table based on the data range
        PivotTableCollection pivotTables = worksheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:B4", "E5", "PivotTable1");
        PivotTable pivotTable = pivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Configure ODS save options (optional settings)
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice   // Set generator type for compatibility
        };

        // Save the workbook with the pivot table to ODS format
        workbook.Save("PivotTableOutput.ods", saveOptions);
    }
}