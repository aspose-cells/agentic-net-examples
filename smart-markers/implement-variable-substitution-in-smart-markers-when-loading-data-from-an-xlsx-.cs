using System;
using Aspose.Cells;
using System.Data;

class SmartMarkerVariableSubstitution
{
    static void Main()
    {
        // Load the template workbook that contains smart markers and a variables sheet
        Workbook workbook = new Workbook("TemplateWithVariables.xlsx");

        // Initialize the WorkbookDesigner and assign the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Specify the name of the worksheet that holds variable smart markers (e.g., &=$VariableName)
        designer.VariablesWorksheetName = "Variables";

        // Set values for the variables defined in the Variables worksheet
        // These variables are referenced in the template using the syntax &=$VariableName
        designer.SetDataSource("VariableName", "John Doe");
        designer.SetDataSource("Age", 30);
        designer.SetDataSource("Date", DateTime.Today);

        // Process all smart markers, including the variable substitutions
        designer.Process();

        // Save the processed workbook
        workbook.Save("Result.xlsx");
    }
}