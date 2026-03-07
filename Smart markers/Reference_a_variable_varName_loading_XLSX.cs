using System;
using Aspose.Cells;

namespace AsposeCellsVariableReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will contain the placeholder
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].PutValue("&=$MyVariable"); // placeholder using smart marker syntax

            // Add a worksheet that will be used for variables (can be empty)
            workbook.Worksheets.Add("Variables");

            // Name of the variable to reference (e.g., {{MyVariable}})
            string varName = "MyVariable";

            // Value to assign to the variable
            string varValue = "Hello from Aspose.Cells!";

            // Create a WorkbookDesigner instance for processing smart markers / variables
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Specify the worksheet that contains the custom variables.
            designer.VariablesWorksheetName = "Variables";

            // Assign a value to the variable. The variable can be used in the template
            // with a smart marker like &=$MyVariable.
            designer.SetDataSource(varName, varValue);

            // Process the smart markers and replace the variable placeholders with the provided value.
            designer.Process();

            // Save the modified workbook to a new XLSX file.
            string outputPath = "ResultWithVariableReplaced.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Variable '{varName}' has been replaced with value '{varValue}'.");
            Console.WriteLine($"Result saved to: {outputPath}");
        }
    }
}