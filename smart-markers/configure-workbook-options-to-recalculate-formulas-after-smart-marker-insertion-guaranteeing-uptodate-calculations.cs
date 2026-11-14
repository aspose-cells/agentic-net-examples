// Title: How to recalculate Excel formulas after inserting smart markers with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# recalculate workbook formulas after WorkbookDesigner.Process | Update Excel calculations after smart marker processing using Aspose.Cells | C# example to call CalculateFormula after smart markers insertion | How to force formula evaluation after using WorkbookDesigner with a DataSet | Ensure formulas are refreshed in generated Excel file with Aspose.Cells
// Tags: Aspose.Cells WorkbookDesigner calculate formulas | C# recalculate Excel after smart marker processing | post‑processing formula evaluation Aspose.Cells | Excel workbook formula refresh with WorkbookDesigner | smart markers workbook.CalculateFormula usage

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

// Loads a template workbook, sets a DataSet as the data source for smart markers, processes them with WorkbookDesigner, calls workbook.CalculateFormula to update all formulas, and saves the resulting Excel file.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";

            // Verify template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Prepare data source for smart markers using a DataSet
            DataSet dataSet = new DataSet();
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Salary", typeof(double));

            // Add sample rows
            employeeTable.Rows.Add("John", 5000);
            employeeTable.Rows.Add("Jane", 6000);

            dataSet.Tables.Add(employeeTable);

            // Process smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Ensure formulas are calculated after processing
            workbook.CalculateFormula();

            // Save the updated workbook
            const string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
