using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a template workbook containing a smart‑marker
        //    with the FORMULA parameter.
        // -------------------------------------------------
        Workbook template = new Workbook();                     // create workbook
        Worksheet sheet = template.Worksheets[0];

        // Place a smart‑marker in A1. The FORMULA parameter will set a
        // formula (=SUM(B1:B3)) for the cell after processing.
        sheet.Cells["A1"].PutValue("${Item:Formula=SUM(B1:B3)}");

        // Add sample numeric data that the formula will sum.
        sheet.Cells["B1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["B3"].PutValue(30);

        // Save the template (optional, shows the intermediate file).
        template.Save("Template.xlsx");

        // -------------------------------------------------
        // 2. Prepare a data source for the smart‑marker.
        // -------------------------------------------------
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Item", typeof(string));
        dt.Rows.Add("Total");   // value that will replace the smart‑marker

        // -------------------------------------------------
        // 3. Load the template with WorkbookDesigner, bind the data,
        //    enable formula calculation and process the smart‑marker.
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook("Template.xlsx");   // load workbook
        designer.SetDataSource(dt);                         // bind data source
        designer.CalculateFormula = true;                   // calculate formulas after processing
        designer.Process();                                 // process smart‑markers

        // -------------------------------------------------
        // 4. Save the resulting workbook.
        // -------------------------------------------------
        designer.Workbook.Save("Result.xlsx");
    }
}