// Title: Calculate employee age with a smart‑marker formula using DATEDIF in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a smart‑marker to an Excel cell which computes age from a BirthDate column using the DATEDIF function and TODAY(). | Show how to bind a DataTable to a WorkbookDesigner, process the smart markers, and save the workbook with the calculated age column.
// Common Searches: Aspose.Cells C# smart marker formula to compute age from birthdate | using DATEDIF with smart markers in Aspose.Cells workbook | bind DataTable to WorkbookDesigner and calculate age column automatically | how to add dynamic age column in Excel using Aspose.Cells smart markers | smart marker date functions example Aspose.Cells .NET
// Tags: smart marker age calculation DATEDIF | Aspose.Cells WorkbookDesigner bind DataTable | C# smart marker date functions | dynamic age column Excel Aspose.Cells | formula parameter date functions Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, defines smart markers for Name, BirthDate, and Age columns, and uses a DATEDIF formula to compute age from the BirthDate smart marker. It binds a DataTable with employee data to the workbook via WorkbookDesigner, processes the markers, and saves the result as SmartMarker_Age_Output.xlsx.
class SmartMarkerAgeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Employees";

        // Set up headers with smart markers
        // Name column will be filled from the data source
        sheet.Cells["A1"].PutValue("&=Name");
        // BirthDate column will be filled from the data source
        sheet.Cells["B1"].PutValue("&=BirthDate");
        // Age column will be calculated using a formula that references the BirthDate smart marker
        // The formula uses DATEDIF to compute the number of full years between BirthDate and today
        sheet.Cells["C1"].PutValue("&=Age: =DATEDIF(&=BirthDate, TODAY(), \"Y\")");

        // Prepare sample data source
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("BirthDate", typeof(DateTime));

        dt.Rows.Add("Alice", new DateTime(1990, 5, 15));
        dt.Rows.Add("Bob", new DateTime(1985, 12, 3));
        dt.Rows.Add("Charlie", new DateTime(2000, 8, 22));

        // Bind the data source to the workbook using WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // Process smart markers

        // Save the result
        workbook.Save("SmartMarker_Age_Output.xlsx");
    }
}
