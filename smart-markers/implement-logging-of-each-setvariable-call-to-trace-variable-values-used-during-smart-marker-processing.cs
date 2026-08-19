// Title: C# Example: Log Every SetVariable Call While Processing Aspose.Cells Smart Markers
// Description: Demonstrates a helper method that writes the variable name and value to the console before invoking WorkbookDesigner.SetDataSource, then creates a template workbook with a smart‑marker, assigns variables, processes the markers, and saves the result. The pattern can be extended to file or database logging for debugging and audit purposes.
// Keywords: Aspose.Cells | C# | .NET | smart markers | WorkbookDesigner | SetDataSource | SetVariable | logging | debugging | variable tracing | console output | code example | GitHub sample
// Common Searches: Aspose.Cells log SetDataSource calls | debug smart marker variables C# | trace variable values in Aspose.Cells | how to log SetVariable in .NET | smart marker debugging example
// Developer Intent: Add runtime logging that records each variable name and its value when it is supplied to WorkbookDesigner.SetDataSource during smart‑marker processing.
// Use Cases: Verify that the correct data is bound to smart markers before generation. | Create an audit trail of all variables used in a report for compliance or troubleshooting. | Switch from console logging to a structured logger (e.g., NLog, Serilog) without changing business logic.
// AI Prompts: Generate a C# extension method for WorkbookDesigner that logs variable assignments to a file and then calls SetDataSource. | Provide code that captures SetVariable calls into a JSON log file while still processing smart markers. | Refactor the SetVariable helper to use Serilog for asynchronous logging of variable names and values.

using System;
using Aspose.Cells;

namespace AsposeCellsVariableLoggingDemo
{
    // Demonstrates a helper method that writes the variable name and value to the console before invoking WorkbookDesigner.SetDataSource, then creates a template workbook with a smart‑marker, assigns variables, processes the markers, and saves the result. The pattern can be extended to file or database logging for debugging and audit purposes.
    class Program
    {
        // Helper method that logs the variable name and value before setting it as a data source.
        static void SetVariable(WorkbookDesigner designer, string variableName, object value)
        {
            // Log the variable assignment.
            Console.WriteLine($"SetVariable called - Name: \"{variableName}\", Value: \"{value}\"");

            // Set the variable (smart marker data source) on the designer.
            designer.SetDataSource(variableName, value);
        }

        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook that will act as the template.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // Add a worksheet that will contain the smart marker referencing a variable.
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";

            // Place a smart marker that uses a variable named "ReportTitle".
            // The syntax "&=$VariableName" tells Aspose.Cells to replace it with the variable's value.
            templateSheet.Cells["A1"].PutValue("&=$ReportTitle");

            // -----------------------------------------------------------------
            // 2. Initialize WorkbookDesigner with the template workbook.
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // -----------------------------------------------------------------
            // 3. Set variables using the logging helper.
            // -----------------------------------------------------------------
            SetVariable(designer, "ReportTitle", "Quarterly Sales Report");

            // You can set additional variables in the same way.
            SetVariable(designer, "GeneratedOn", DateTime.Now);

            // -----------------------------------------------------------------
            // 4. Process the smart markers – variables will be replaced with the logged values.
            // -----------------------------------------------------------------
            designer.Process();

            // -----------------------------------------------------------------
            // 5. Save the resulting workbook.
            // -----------------------------------------------------------------
            // The save operation follows the standard Aspose.Cells pattern.
            workbook.Save("VariableLoggingResult.xlsx");

            Console.WriteLine("Workbook saved as VariableLoggingResult.xlsx");
        }
    }
}
