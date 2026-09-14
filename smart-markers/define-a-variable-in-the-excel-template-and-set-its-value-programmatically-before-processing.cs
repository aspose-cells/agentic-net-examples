// Title: How to set a smart‑marker variable value programmatically in an Aspose.Cells Excel template (C#)
// AI Prompts: Create an Excel template containing a smart‑marker variable and assign its value using WorkbookDesigner.SetDataSource in C#. | Replace the &=$MyVar smart marker with a custom string before processing the workbook via code. | Show how to use a separate variables worksheet with WorkbookDesigner to bind variable data for smart markers in Aspose.Cells.
// Common Searches: set smart marker variable value Aspose.Cells C# example | WorkbookDesigner SetDataSource usage for Excel templates | replace &=$MyVar smart marker with custom text programmatically | define variables worksheet for smart markers Aspose.Cells C#
// Tags: WorkbookDesigner SetDataSource for smart marker variable C# | smart marker variable worksheet Aspose.Cells | dynamic variable binding in Excel template C# | Aspose.Cells programmatic variable replacement | C# generate Excel with smart marker variable

using System;
using Aspose.Cells;

namespace AsposeCellsVariableDemo
{
    // Demonstrates creating a workbook, adding a template sheet with a smart‑marker variable, optionally defining a variables sheet, using WorkbookDesigner to bind the variable value via SetDataSource, processing the smart markers, and saving the resulting Excel file.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // 2. Add a worksheet that will contain the smart marker for the variable
            Worksheet templateSheet = workbook.Worksheets.Add("Template");
            // Place a smart marker that references a variable named "MyVar"
            // The syntax "&=$VariableName" tells WorkbookDesigner to replace it with the variable value
            templateSheet.Cells["A1"].PutValue("&=$MyVar");

            // 3. (Optional) Add a worksheet that could hold variable definitions.
            // Not required when using SetDataSource for a variable, but shown for completeness.
            Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
            // Example layout: Column A = variable name, Column B = value
            variablesSheet.Cells["A1"].PutValue("MyVar");
            variablesSheet.Cells["B1"].PutValue("Placeholder"); // will be overwritten programmatically

            // 4. Initialize WorkbookDesigner with the workbook (lifecycle rule: create)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // 5. Specify the worksheet that contains variable definitions (if you rely on that sheet)
            designer.VariablesWorksheetName = "Variables";

            // 6. Set the variable value programmatically.
            // This binds the variable name "MyVar" to the desired value.
            designer.SetDataSource("MyVar", "Hello from Aspose.Cells!");

            // 7. Process the smart markers so the variable value replaces the placeholder
            designer.Process();

            // 8. Save the resulting workbook (lifecycle rule: save)
            workbook.Save("VariableDemo_Output.xlsx");

            Console.WriteLine("Workbook created and variable value set successfully.");
        }
    }
}
