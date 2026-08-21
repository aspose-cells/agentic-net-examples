// Title: Inject Runtime Variables into Aspose.Cells Smart Markers with WorkbookDesigner.SetDataSource (C#)
// Description: Demonstrates how to create a workbook, add a variables sheet, place smart markers that reference a runtime variable, assign the variable value using WorkbookDesigner.SetDataSource, process the markers, and save the result. Shows the variable being used directly in a cell value and inside a formula.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | smart markers | runtime variable injection | C# Excel automation | dynamic calculations | Excel template variables | discount smart marker | variables worksheet
// Common Searches: Aspose.Cells set variable for smart markers | C# inject runtime value into Excel smart marker | WorkbookDesigner SetDataSource example | how to use variables worksheet in Aspose.Cells | dynamic discount calculation with smart markers
// Developer Intent: Assign a value to a smart‑marker variable at runtime so that the marker and any formulas referencing it are evaluated with the supplied data.
// Use Cases: Replace a placeholder smart marker with a discount rate and apply it in a calculation formula. | Maintain a dedicated "Variables" worksheet, populate multiple parameters (tax, commission, etc.) via SetDataSource, and generate a report with consistent values across sheets. | Create reusable Excel templates where a single variable (e.g., exchange rate) is injected once and automatically propagated to all smart‑marker expressions.
// AI Prompts: Generate C# code that uses WorkbookDesigner.SetDataSource to inject a "TaxRate" variable into smart markers and recalculate dependent formulas. | Explain the steps to configure VariablesWorksheetName, set several variables, and process smart markers in Aspose.Cells. | Show how to verify that smart‑marker expressions using injected variables produce the expected numeric results after processing.

using System;
using Aspose.Cells;

namespace AsposeCellsVariableInjectionDemo
{
    // Demonstrates how to create a workbook, add a variables sheet, place smart markers that reference a runtime variable, assign the variable value using WorkbookDesigner.SetDataSource, process the markers, and save the result. Shows the variable being used directly in a cell value and inside a formula.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Add a worksheet that will hold variable definitions (optional, but we set the name)
                Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
                // Placeholder for demonstration (not required for SetDataSource variables)
                variablesSheet.Cells["A1"].PutValue("Variable Definitions");

                // 3. Add a template worksheet that contains smart markers referencing variables
                Worksheet templateSheet = workbook.Worksheets.Add("Template");
                // Smart marker that will be replaced by the variable value
                templateSheet.Cells["A1"].PutValue("&=$Discount");
                // Use the variable in a formula via a smart marker expression
                templateSheet.Cells["A2"].Formula = "=100*(&=$Discount)";

                // 4. Create a WorkbookDesigner and associate it with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // 5. Specify the worksheet that contains variables (optional, shown for completeness)
                designer.VariablesWorksheetName = "Variables";

                // 6. Inject runtime variable values using SetDataSource(string, object)
                designer.SetDataSource("Discount", 0.15); // 15% discount

                // 7. Process the smart markers so that they are replaced with the injected values
                designer.Process();

                // 8. Save the resulting workbook
                workbook.Save("VariableInjectionResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
