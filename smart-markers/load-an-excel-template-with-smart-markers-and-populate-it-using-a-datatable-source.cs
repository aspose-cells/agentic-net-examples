// Title: Populate an Excel template using smart markers from a C# DataTable with Aspose.Cells WorkbookDesigner
// AI Prompts: Write C# code that loads an existing Excel file containing smart markers, creates a DataTable with employee data, binds the DataTable to the "Employees" smart‑marker group via WorkbookDesigner, processes the markers, and saves the populated workbook. | Show how to detect a missing template file, create a new workbook on the fly, then apply smart markers using a DataTable source with Aspose.Cells. | Provide an example of robust error handling while populating an Excel template with smart markers from a DataTable in C#.
// Common Searches: Aspose.Cells C# example for filling smart markers from a DataTable | How to bind a DataTable to a smart‑marker group in an Excel template using WorkbookDesigner | Populate Excel template with employee data using Aspose.Cells smart markers | C# code to process smart markers in a workbook and save the result | Create Excel file from template when the template file is missing Aspose.Cells
// Tags: Aspose.Cells WorkbookDesigner smart‑marker population | C# DataTable binding to Excel smart markers | Excel template processing with Aspose.Cells | populate workbook from DataTable Aspose.Cells | handle missing Excel template Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel template (or creates a new workbook if the file is absent), builds a DataTable with employee information, assigns the table to the "Employees" smart‑marker group via WorkbookDesigner, processes the smart markers to fill the template, and saves the resulting workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string templatePath = "Template.xlsx";
            Workbook workbook;

            // Load the template if it exists; otherwise create a new workbook
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Initialize WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Create and fill a DataTable for the smart marker group "Employees"
            DataTable dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Salary", typeof(double));

            dt.Rows.Add("John Doe", 30, 50000);
            dt.Rows.Add("Jane Smith", 28, 60000);
            dt.Rows.Add("Bob Johnson", 35, 55000);

            // Assign the DataTable to the smart marker group
            designer.SetDataSource("Employees", dt);

            // Process smart markers and populate the workbook
            designer.Process();

            // Save the populated workbook
            string resultPath = "Result.xlsx";
            designer.Workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
