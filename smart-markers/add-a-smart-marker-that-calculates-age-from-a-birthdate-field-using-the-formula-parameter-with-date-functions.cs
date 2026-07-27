using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerAgeDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create template workbook ----------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("BirthDate");
            cells["B1"].PutValue("Age");

            // Smart marker for the birthdate value
            cells["A2"].PutValue("&=Employees.BirthDate");

            // Smart marker with a formula that calculates age using DATEDIF
            // The smart marker syntax for a formula is: &=\"=YourFormula\"
            // Inside the formula we reference the birthdate smart marker again.
            string ageFormulaSmartMarker = "&=\"=DATEDIF(&=Employees.BirthDate, TODAY(), \\\"Y\\\")\"";
            cells["B2"].PutValue(ageFormulaSmartMarker);

            // ---------- Prepare data source ----------
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("BirthDate", typeof(DateTime));

            // Sample birth dates
            employees.Rows.Add(DateTime.Now.AddYears(-30));               // 30 years old
            employees.Rows.Add(DateTime.Now.AddYears(-25).AddMonths(-3)); // 25 years old (approx)
            employees.Rows.Add(DateTime.Now.AddYears(-45).AddDays(-10)); // 45 years old (approx)

            // ---------- Process smart markers ----------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = template;          // assign the template workbook
            designer.SetDataSource(employees);     // bind the data source
            designer.CalculateFormula = true;      // ensure formulas are calculated after binding
            designer.Process();                    // populate smart markers and calculate formulas

            // ---------- Save result ----------
            designer.Workbook.Save("AgeSmartMarkerOutput.xlsx");
        }
    }
}