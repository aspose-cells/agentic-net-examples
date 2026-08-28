// Title: Create a C# method that loads an Excel template, binds any data source with Aspose.Cells smart markers, and returns the populated workbook as a MemoryStream (XLS)
// AI Prompts: Write a C# function that receives an object (list, DataTable, or custom class) and a file path to an Excel template, uses WorkbookDesigner to set the data source named "Data", processes the smart markers, and returns the result as a MemoryStream. | Show how to invoke the function with a List<Person>, obtain the MemoryStream, and write it to an .xls file on disk. | Explain how to validate the template file path before creating the WorkbookDesigner to avoid FileNotFoundException. | Demonstrate returning the generated MemoryStream from a Web API endpoint for client download.
// Common Searches: asp.net core return excel file from template using aspose.cells smart markers | c# generate memorystream from excel template with workbookdesigner | how to bind a list of objects to aspose.cells smart markers in a template | save processed aspose.cells workbook to stream instead of file | populate excel template with custom object and get xls stream in .net
// Tags: aspose.cells workbookdesigner data binding | excel template to memorystream c# | smart markers generate xls stream | c# populate excel template using aspose.cells | return workbook as memorystream after processing markers

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel template, assigns the supplied data object to the smart marker name "Data" via WorkbookDesigner, processes the markers, and returns the filled workbook as a MemoryStream (XLS), which can then be saved to a file or sent to a client.
public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Determine template path (default to "Template.xlsx" if not provided)
            string templatePath = args.Length > 0 ? args[0] : "Template.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Example data source: a list of Person objects
            var data = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 25 }
            };

            // Generate the workbook using the utility method
            MemoryStream workbookStream = WorkbookGenerator.GenerateWorkbook(data, templatePath);

            // Save the generated workbook to an output file
            string outputPath = "Output.xls";
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                workbookStream.WriteTo(fileStream);
            }

            Console.WriteLine($"Workbook generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Simple data class used in the example
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}

public static class WorkbookGenerator
{
    /// <param name="data">The data source object (e.g., a list, DataTable, or custom class).</param>
    /// <param name="templatePath">Full path to the Excel template containing smart markers.</param>
    /// <returns>A MemoryStream containing the populated workbook (XLS format).</returns>
    public static MemoryStream GenerateWorkbook(object data, string templatePath)
    {
        try
        {
            // Load the template workbook from the specified file path.
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner and associate it with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the provided data object to a smart marker name (e.g., "Data").
            // The template should contain markers like &Data.PropertyName.
            designer.SetDataSource("Data", data);

            // Process the smart markers to populate the worksheet with data.
            designer.Process();

            // Save the processed workbook to a memory stream (XLS format) and return it.
            return workbook.SaveToStream();
        }
        catch
        {
            // Rethrow the exception to be handled by the caller.
            throw;
        }
    }
}
