// Title: Bind multiple DataSet tables to distinct smart‑marker groups using WorkbookDesigner.SetDataSource in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a WorkbookDesigner, sets a DataSet with Employees and Departments tables using SetDataSource, runs the smart‑marker engine, and writes the workbook to disk. | Show how to place smart‑marker placeholders for two groups (Employees and Departments) in a worksheet, associate each group with its matching DataTable, and calculate any formulas after the engine runs. | Provide a robust example with try‑catch that builds the DataSet, binds it, invokes WorkbookDesigner.Process, calls Workbook.CalculateFormula, and saves the resulting Excel file.
// Common Searches: Aspose.Cells C# bind DataSet with multiple tables to different smart marker groups | WorkbookDesigner SetDataSource example with Employees and Departments tables | How to run smart markers for several DataTables in one Excel workbook using Aspose.Cells | C# generate Excel file with smart markers from a DataSet containing two tables | Calculate formulas after smart marker processing Aspose.Cells .NET
// Tags: WorkbookDesigner SetDataSource multiple DataTables | smart markers separate groups Employees Departments | Aspose.Cells generate Excel from DataSet | C# bind DataSet to smart marker groups | calculate formulas after smart marker processing

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkersExample
{
    // The example creates a workbook, defines two smart‑marker groups (Employees and Departments), builds a DataSet with matching tables, binds the DataSet using WorkbookDesigner.SetDataSource, processes the markers, calculates any formulas, and saves the file as SmartMarkersOutput.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define smart markers for the Employees group
                sheet.Cells["A1"].PutValue("&=Employees.Name");
                sheet.Cells["B1"].PutValue("&=Employees.Salary");

                // Define smart markers for the Departments group
                sheet.Cells["A5"].PutValue("&=Departments.DeptName");
                sheet.Cells["B5"].PutValue("&=Departments.Location");

                // Build a DataSet with two tables: Employees and Departments
                DataSet ds = new DataSet();

                // Employees table
                DataTable employees = new DataTable("Employees");
                employees.Columns.Add("Name", typeof(string));
                employees.Columns.Add("Salary", typeof(double));
                employees.Rows.Add("John", 50000);
                employees.Rows.Add("Jane", 60000);
                ds.Tables.Add(employees);

                // Departments table
                DataTable departments = new DataTable("Departments");
                departments.Columns.Add("DeptName", typeof(string));
                departments.Columns.Add("Location", typeof(string));
                departments.Rows.Add("HR", "New York");
                departments.Rows.Add("IT", "San Francisco");
                ds.Tables.Add(departments);

                // Process smart markers using WorkbookDesigner (correct API usage)
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = workbook;
                designer.SetDataSource(ds);
                designer.Process();

                // Calculate any formulas that may exist
                workbook.CalculateFormula();

                // Prepare output path and ensure directory exists
                string outputPath = "SmartMarkersOutput.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
