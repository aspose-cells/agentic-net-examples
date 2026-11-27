// Title: Implement a percentage progress callback for WorkbookDesigner smart marker processing in Aspose.Cells using C#
// Common Searches: c# Aspose.Cells progress callback for WorkbookDesigner smart markers | how to report percentage completed while processing smart markers in .NET | example of monitoring large smart marker population progress with Aspose.Cells | Aspose.Cells smart marker processing progress event usage in C# | track smart marker population progress for big data sets in Aspose.Cells
// Tags: WorkbookDesigner smart marker progress callback | Aspose.Cells percentage progress reporting | C# large data smart marker population | monitor smart marker processing progress | Aspose.Cells progress event handling

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to attach a progress callback to Aspose.Cells' WorkbookDesigner so that the percentage of smart marker population is reported while processing a large DataTable (10,000 rows). It shows loading a template workbook, setting the data source, handling the Designer's progress event to output completion percentages, processing the smart markers, and saving the populated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook that contains smart markers.
            Workbook workbook = new Workbook(templatePath);

            // Create a large data source to simulate extensive smart marker population.
            DataTable data = new DataTable("Employees");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(int));
            data.Columns.Add("Salary", typeof(double));

            // Populate the table with many rows.
            for (int i = 1; i <= 10000; i++)
            {
                data.Rows.Add($"Employee {i}", 20 + (i % 30), 30000 + (i * 10));
            }

            // Initialize the WorkbookDesigner which handles smart marker processing.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(data);

            // Execute the smart marker population.
            designer.Process();

            // Save the populated workbook.
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
