// Title: Use Aspose.Cells WorkbookDesigner in C# to populate smart markers on hidden worksheets
// AI Prompts: Generate C# code that loads an Excel template, creates a DataSet with employee information, binds it to a WorkbookDesigner, processes all smart markers—including those on hidden worksheets—and saves the output file. | Write a C# example that ensures hidden sheets are included when processing smart markers with Aspose.Cells WorkbookDesigner, handling missing template files and creating the result directory.
// Common Searches: Aspose.Cells WorkbookDesigner process smart markers on hidden worksheets C# | C# populate hidden Excel sheet using smart markers Aspose.Cells | How to include hidden tabs when using smart markers with Aspose.Cells .NET | Example of binding a DataSet to WorkbookDesigner for hidden sheet data population
// Tags: WorkbookDesigner hidden worksheet processing | smart markers hidden sheet population | Aspose.Cells DataSet binding for hidden tabs | C# Aspose.Cells populate hidden worksheets | Excel template smart markers hidden sheets

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads a template workbook, builds a DataSet containing an Employees table, assigns it to a WorkbookDesigner, processes smart markers on every sheet—including hidden ones—and saves the populated workbook to Result.xlsx while handling missing files and ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: The template file \"{templatePath}\" was not found.");
                return;
            }

            // Load the workbook that contains hidden worksheets with smart markers
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // ----- Prepare the data source -----
            DataSet dataSet = new DataSet();

            DataTable employeesTable = new DataTable("Employees");
            employeesTable.Columns.Add("Name", typeof(string));
            employeesTable.Columns.Add("Age", typeof(int));
            employeesTable.Columns.Add("Department", typeof(string));

            // Add sample rows (replace with real data as needed)
            employeesTable.Rows.Add("John Doe", 30, "Sales");
            employeesTable.Rows.Add("Jane Smith", 28, "Marketing");
            employeesTable.Rows.Add("Bob Johnson", 35, "IT");

            dataSet.Tables.Add(employeesTable);

            // Assign the data source to the designer
            designer.SetDataSource(dataSet);

            // Process all smart markers, including those located in hidden worksheets
            designer.Process();

            // Ensure the directory for the result file exists
            string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
            {
                Directory.CreateDirectory(resultDir);
            }

            // Save the populated workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
