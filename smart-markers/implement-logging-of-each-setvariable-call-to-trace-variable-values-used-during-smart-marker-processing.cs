// Title: Log each WorkbookDesigner.SetDataSource call to trace smart marker variable values in Aspose.Cells for .NET
// Common Searches: how to log WorkbookDesigner SetDataSource values when using Aspose.Cells smart markers in C# | debug smart marker variable assignments in Aspose.Cells .NET | trace variable values passed to Aspose.Cells smart markers during processing | log smart marker data source values to console with Aspose.Cells | C# example for logging smart marker variables in Aspose.Cells workbook designer
// Tags: WorkbookDesigner SetDataSource trace | Aspose.Cells smart marker variable logging | C# smart marker debugging | console logging for Aspose.Cells variables | trace data source assignments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Provides a SetVariableLogged helper that formats each variable (including nulls and arrays), writes a detailed console message with name, readable value, and full type, then binds the variable using WorkbookDesigner.SetDataSource, enabling step‑by‑step tracing of smart marker values during Aspose.Cells processing.
class Program
{
    static void Main()
    {
        // Paths to the template and output files (place the template in the same folder as the executable)
        string templatePath = "Template.xlsx";
        string outputPath = "Result.xlsx";

        // Verify that the template file exists
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Error: Template file not found at '{Path.GetFullPath(templatePath)}'.");
            return;
        }

        try
        {
            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Create a WorkbookDesigner instance for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Example variables – each call is logged via SetVariableLogged
            SetVariableLogged(designer, "CustomerName", "John Doe");
            SetVariableLogged(designer, "OrderDate", DateTime.Now);
            SetVariableLogged(designer, "TotalAmount", 1234.56);
            SetVariableLogged(designer, "IsPremium", true);
            SetVariableLogged(designer, "Items", new string[] { "Item1", "Item2", "Item3" });

            // Process all smart markers in the workbook
            designer.Process();

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Output any unexpected errors
            Console.WriteLine($"Exception: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    /// <param name="designer">The WorkbookDesigner instance.</param>
    /// <param name="name">The variable name used in smart markers.</param>
    /// <param name="value">The value to assign to the variable.</param>
    static void SetVariableLogged(WorkbookDesigner designer, string name, object value)
    {
        // Prepare a readable representation of the value
        string valueRepresentation;
        if (value == null)
        {
            valueRepresentation = "null";
        }
        else if (value is Array arr)
        {
            // For arrays, show the type and length
            valueRepresentation = $"Array of {value.GetType().GetElementType().Name}[{arr.Length}]";
        }
        else
        {
            valueRepresentation = value.ToString();
        }

        // Log to console (could be extended to log to a file or other logging framework)
        Console.WriteLine($"[SetVariable] Name: '{name}', Value: '{valueRepresentation}', Type: {value?.GetType().FullName ?? "null"}");

        // Bind the variable to the designer using the correct API
        designer.SetDataSource(name, value);
    }
}
