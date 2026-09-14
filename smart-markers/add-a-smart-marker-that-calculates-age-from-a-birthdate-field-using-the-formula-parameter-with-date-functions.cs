// Title: Use a smart marker with the Formula parameter to compute age from a birthdate in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a smart marker to an Aspose.Cells template which calculates age using the DATEDIF and TODAY functions. | Show how to bind a DataTable of employee names and birthdates to WorkbookDesigner, enable formula calculation, and process the smart markers. | Explain how to modify a worksheet cell to include a smart‑marker expression that returns the calculated age in the resulting Excel file.
// Common Searches: aspnet calculate age in Excel using Aspose.Cells smart markers DATEDIF | C# Aspose.Cells smart marker formula parameter example with date functions | how to bind DataTable to WorkbookDesigner and compute age column | smart marker to display years between birthdate and today in Aspose.Cells | enable formula evaluation after processing smart markers Aspose.Cells
// Tags: smart marker formula age calculation | Aspose.Cells DATEDIF function | WorkbookDesigner bind DataTable C# | calculate age from birthdate Aspose.Cells | enable formula evaluation after smart marker processing

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerAgeDemo
{
    // Demonstrates creating an Aspose.Cells template workbook with smart markers for Name, BirthDate, and Age, where the Age cell uses the Formula parameter with DATEDIF and TODAY to compute years. Shows how to populate a DataTable of employees, bind it to WorkbookDesigner, enable formula calculation, process the markers, and save the output file (EmployeesWithAge.xlsx).
    class Program
    {
        static void Main()
        {
            // ---------- Create a template workbook ----------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("BirthDate");
            cells["C1"].PutValue("Age");

            // Smart markers for data rows
            // Name smart marker
            cells["A2"].PutValue("&=Employees.Name");
            // BirthDate smart marker (raw date value)
            cells["B2"].PutValue("&=Employees.BirthDate");
            // Age smart marker using Formula parameter with date functions
            // DATEDIF calculates the difference in years between birthdate and today
            cells["C2"].PutValue("&=Employees.BirthDate?Formula=DATEDIF(&=Employees.BirthDate,TODAY(),\"Y\")");

            // ---------- Prepare data source ----------
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("BirthDate", typeof(DateTime));

            dt.Rows.Add("John Doe", new DateTime(1990, 5, 15));
            dt.Rows.Add("Jane Smith", new DateTime(1985, 12, 3));
            dt.Rows.Add("Bob Johnson", new DateTime(2000, 7, 22));

            // ---------- Process smart markers ----------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = template;               // Load the template workbook
            designer.SetDataSource(dt);                 // Bind the DataTable
            designer.CalculateFormula = true;           // Ensure formulas are calculated after processing
            designer.Process();                         // Populate smart markers

            // ---------- Save the result ----------
            designer.Workbook.Save("EmployeesWithAge.xlsx");
        }
    }
}
