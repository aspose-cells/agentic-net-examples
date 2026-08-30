// Title: Save a workbook with a pivot table as an OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates sample data, builds a pivot table, sets OdsSaveOptions with the LibreOffice generator, and saves the workbook to an .ods file using Aspose.Cells. | Show how to add a PivotTable to a worksheet and export the entire workbook to OpenDocument format with the Aspose.Cells .NET API. | Demonstrate configuring OdsSaveOptions.GeneratorType for LibreOffice when saving a workbook that contains a pivot table as ODS.
// Common Searches: Aspose.Cells .NET how to export a workbook with a pivot table to ODS | C# save Excel file as OpenDocument spreadsheet including pivot tables | Configure OdsSaveOptions generator type LibreOffice in Aspose.Cells example | Create pivot table programmatically and convert to .ods using Aspose.Cells | Save workbook as ODS format with pivot data using Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table ODS conversion | OdsSaveOptions generator type setting | C# create and export pivot table to OpenDocument | Aspose.Cells save workbook as .ods | programmatic pivot table generation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

namespace AsposeCellsPivotToOds
{
    // The example creates a new workbook, fills it with product and sales data, adds a pivot table summarizing the sales, configures OdsSaveOptions to use the LibreOffice generator, and saves the workbook—including the pivot table—as an OpenDocument Spreadsheet file named PivotTable.ods.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(3000);

            // Create a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:B4", "E5", "PivotTable1");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

            // Configure ODS save options (optional: set generator type)
            OdsSaveOptions odsOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the workbook (including the pivot table) as an ODS file
            workbook.Save("PivotTable.ods", odsOptions);
        }
    }
}
