// Title: Record each successful Smart Marker merge with a custom logger while processing Aspose.Cells templates in C#
// Common Searches: C# Aspose.Cells log smart marker merge events | how to capture smart marker processing notifications in Aspose.Cells .NET | custom logger for Aspose.Cells WorkbookDesigner smart markers | record each successful smart marker merge using Aspose.Cells API | Aspose.Cells smart marker processing with custom logging example
// Tags: Aspose.Cells WorkbookDesigner smart marker logging | custom logger Aspose.Cells smart marker merge | C# Aspose.Cells smart marker processing notification | smart marker merge event logging .NET | Aspose.Cells template processing with custom logger

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, creates a DataSet with employee data, assigns it to a WorkbookDesigner, processes smart markers, logs a success message through a simple CustomLogger, and saves the resulting workbook while handling missing files and runtime exceptions.
class CustomLogger
{
    public void Log(string message)
    {
        // Simple logging to console (could be extended to file, etc.)
        Console.WriteLine(message);
    }
}

class Program
{
    static void Main()
    {
        const string templatePath = "Template.xlsx";
        const string resultPath = "Result.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        try
        {
            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Prepare the data source for smart markers using a DataSet
            DataSet dataSet = new DataSet();
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John", 30);
            employeeTable.Rows.Add("Jane", 25);

            dataSet.Tables.Add(employeeTable);

            // Attach the data source to the designer
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);

            // Instantiate the custom logger
            CustomLogger logger = new CustomLogger();

            // Process smart markers
            designer.Process();

            // Log that processing completed
            logger.Log("Smart markers processed successfully.");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the resulting workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
