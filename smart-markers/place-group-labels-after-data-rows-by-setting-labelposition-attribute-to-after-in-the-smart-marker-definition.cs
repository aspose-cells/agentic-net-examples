// Title: Set smart marker group label to appear after data rows with Aspose.Cells in C#
// AI Prompts: Write C# code that opens an Excel template, finds the smart marker group definition, sets its LabelPosition attribute to "After", processes a DataSet with employee data using WorkbookDesigner, and saves the result workbook. | Show how to edit the smart marker XML inside a workbook to move the group label to the bottom of the generated rows using Aspose.Cells.
// Common Searches: Aspose.Cells C# set smart marker group label after rows | change smart marker LabelPosition to After in Excel template | how to place group labels after data rows with Aspose.Cells | modify smart marker XML labelposition attribute Aspose.Cells | C# Aspose.Cells group label placement after generated rows
// Tags: Aspose.Cells set smart marker labelposition | C# edit smart marker XML | group label after rows Aspose.Cells | WorkbookDesigner process smart markers with after label | Excel template smart marker label placement

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

// This example demonstrates loading an Excel template, updating the smart marker definition to set LabelPosition="After" so that group labels appear after the data rows, binding a DataSet containing employee information, processing the smart markers with WorkbookDesigner, and saving the resulting workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";

            // Verify that the template file exists before attempting to load it
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook that contains the smart marker definitions
            Workbook workbook = new Workbook(templatePath);

            // Prepare data for the smart markers using a DataSet (required by WorkbookDesigner)
            DataSet dataSet = new DataSet();

            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John", 30);
            employeeTable.Rows.Add("Jane", 25);

            dataSet.Tables.Add(employeeTable);

            // Create a WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the data source and process the smart markers
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the resulting workbook
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
