// Title: How to add a console progress callback that shows percentage while processing large smart markers with Aspose.Cells WorkbookDesigner in C#
// AI Prompts: Write C# code that registers a custom IProgress<double> handler to receive percentage updates from WorkbookDesigner during smart marker expansion. | Show how to display a console progress bar reflecting the smart marker processing progress when populating a large DataSet with Aspose.Cells. | Create a reusable method that accepts a WorkbookDesigner instance and an IProgress<int> to report smart marker population status for any dataset.
// Common Searches: Aspose.Cells how to monitor smart marker processing progress in C# | C# console progress bar for WorkbookDesigner smart marker fill | report percentage while populating Excel smart markers with large DataSet
// Tags: WorkbookDesigner progress callback | Aspose.Cells smart marker progress | C# smart marker processing status | Excel smart marker fill progress | large dataset smart marker reporting

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// This example loads an existing Excel workbook, creates a DataSet with employee information, assigns it as the data source for smart markers via WorkbookDesigner, attaches a progress callback to report percentage completion during the smart marker processing, executes the population, and saves the resulting workbook, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(inputPath);

            // Prepare data source for smart markers using a DataSet
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John", 30);
            employeeTable.Rows.Add("Jane", 25);
            // Add more rows as needed for large‑scale population

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(employeeTable);

            // Initialize the WorkbookDesigner for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the data source for the smart markers
            designer.SetDataSource(dataSet);

            // Process the smart markers
            designer.Process();

            // Save the populated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
