// Title: How to populate Excel smart markers from a DataTable using WorkbookDesigner.SetDataSource in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx template, creates a DataTable with employee records, binds it to the workbook via WorkbookDesigner.SetDataSource, processes the smart markers, and saves the populated file. | Show the step‑by‑step usage of Aspose.Cells WorkbookDesigner to map DataTable columns to smart marker fields and generate a filled Excel workbook in a .NET application.
// Common Searches: aspnet example using WorkbookDesigner.SetDataSource to fill smart markers from a DataTable | c# populate Excel template smart markers with relational data using Aspose.Cells | how to bind DataTable to smart markers in an Excel file with Aspose.Cells | process smart markers in .xlsx template after setting DataSource in C# | Aspose.Cells tutorial for filling smart markers from database query results
// Tags: WorkbookDesigner data binding DataTable | populate smart markers Aspose.Cells | C# fill Excel template from relational data | process smart markers .NET | Aspose.Cells bind DataTable to smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an Excel template that contains smart markers, creates a DataTable with sample employee data, binds the table to the workbook using WorkbookDesigner.SetDataSource, processes the smart markers to replace them with actual values, and saves the resulting workbook as a new file.
class SmartMarkerExample
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";
            const string outputPath = "OutputPopulated.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the Excel template containing smart markers
            Workbook workbook = new Workbook(templatePath);

            // Create a sample DataTable to simulate database data
            DataTable dt = new DataTable();
            dt.Columns.Add("EmployeeID", typeof(int));
            dt.Columns.Add("FirstName", typeof(string));
            dt.Columns.Add("LastName", typeof(string));
            dt.Columns.Add("Salary", typeof(decimal));

            // Add sample rows
            dt.Rows.Add(1, "John", "Doe", 55000m);
            dt.Rows.Add(2, "Jane", "Smith", 62000m);
            dt.Rows.Add(3, "Bob", "Johnson", 48000m);

            // Bind the DataTable to the smart markers and process them
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the populated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
