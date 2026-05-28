using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

class SavePivotTableAsOds
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

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E5", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field
        pivotTable.SaveData = true; // Ensure pivot data is saved with the file

        // Configure ODS save options
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice; // Set generator (optional)
        saveOptions.IgnorePivotTables = false; // Include pivot tables in the ODS file

        // Save the workbook as an ODS file using the specified options
        workbook.Save("PivotTableOutput.ods", saveOptions);
    }
}