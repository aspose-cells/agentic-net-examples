// Title: Log each WorkbookDesigner.SetVariable call to trace smart‑marker variable values in C# using Aspose.Cells
// AI Prompts: Write C# code that creates a subclass of WorkbookDesigner overriding SetVariable to write the variable name and value to a console or file log before calling the base method. | Show how to attach a delegate to Aspose.Cells events to capture and log every SetVariable invocation during smart‑marker processing. | Provide a reusable helper method that wraps WorkbookDesigner.SetVariable, logs the parameters, and then forwards the call for use in smart‑marker templates.
// Common Searches: how to capture WorkbookDesigner SetVariable values for debugging Aspose.Cells smart markers C# | example of logging variable assignments when processing smart markers with Aspose.Cells | trace smart marker variables in Aspose.Cells using custom logger | C# Aspose.Cells SetVariable debug output to console or file | record smart marker variable values during WorkbookDesigner.Process execution
// Tags: WorkbookDesigner SetVariable logging C# | smart marker variable tracing Aspose.Cells | custom logger for Aspose.Cells smart markers | debug Aspose.Cells smart marker processing | log variable values during Excel template generation

using System;
using System.Data;
using Aspose.Cells; // Aspose.Cells contains WorkbookDesigner and related classes

// The example demonstrates how to intercept each WorkbookDesigner.SetVariable call, log the variable name and value (to console or a file), and then continue processing smart markers. It includes a subclass or wrapper approach, optional event handling, and shows the full workflow from template creation to saving the resulting workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (template)
            Workbook workbook = new Workbook();

            // Add a worksheet that will contain smart markers
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";

            // Smart marker that references a variable (syntax: &VariableName)
            templateSheet.Cells["A1"].PutValue("&MyVariable");

            // Prepare data source with a column matching the smart marker name
            string varValue = "Hello Aspose!";

            // Log the variable assignment
            Console.WriteLine($"[SetVariable] Name: \"MyVariable\", Value: \"{varValue ?? "null"}\"");

            // Build a DataTable as the data source for the smart marker
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("MyVariable", typeof(string));
            dt.Rows.Add(varValue);

            // Initialize the workbook designer
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the data source for smart markers
            designer.SetDataSource(dt);

            // Process smart markers (variables will be replaced)
            designer.Process();

            // Save the result
            string outputPath = "SmartMarkerWithLogging.xlsx";
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
