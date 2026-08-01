// Title: Inject Runtime Variables into Smart Markers with SetVariable (C# Aspose.Cells)
// Description: Demonstrates how to add a "Variables" worksheet, store a discount value, reference it in a smart marker ("&=$Discount"), configure WorkbookDesigner.VariablesWorksheetName, process all smart markers, recalculate formulas, and save the workbook. Ideal for dynamic pricing or any scenario where values must be supplied at runtime.
// Keywords: Aspose.Cells SetVariable | WorkbookDesigner smart markers | C# runtime variable injection | dynamic Excel calculations | smart marker discount example | recalculate formulas Aspose.Cells | Excel template variables .NET | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells SetVariable example C# | how to use WorkbookDesigner with variables worksheet | inject discount into smart marker formula | process smart markers and recalculate formulas | dynamic Excel report generation Aspose.Cells
// Developer Intent: Add a runtime variable to a workbook so smart markers can use it during processing.
// Use Cases: Create a pricing sheet where the discount rate is defined once and applied to multiple calculations via smart markers. | Generate region‑specific financial reports by swapping variable values (tax, commission, exchange rate) without altering the template layout. | Automate invoice generation where promotional codes are stored in a separate sheet and injected into smart markers at runtime.
// AI Prompts: Show C# code that uses WorkbookDesigner.SetVariable to assign a discount and apply it to smart markers before processing. | Give an example of loading variable values from a JSON file and injecting them into smart markers with SetVariable in Aspose.Cells. | Explain how to recalculate all dependent formulas after processing smart markers that reference runtime variables.

using System;
using Aspose.Cells;

// Demonstrates how to add a "Variables" worksheet, store a discount value, reference it in a smart marker ("&=$Discount"), configure WorkbookDesigner.VariablesWorksheetName, process all smart markers, recalculate formulas, and save the workbook. Ideal for dynamic pricing or any scenario where values must be supplied at runtime.
class SetVariableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will contain variable definitions
            Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
            // Header for the variable name
            variablesSheet.Cells["A1"].PutValue("Discount");
            // Set the runtime value for the variable directly in the worksheet
            variablesSheet.Cells["B1"].PutValue(0.2); // 20% discount

            // Add a template worksheet where smart markers will be used
            Worksheet templateSheet = workbook.Worksheets.Add("Template");
            // Example data: a price value
            templateSheet.Cells["A1"].PutValue(100); // Price
            // Smart marker that will be replaced by the variable value at runtime
            templateSheet.Cells["B1"].PutValue("&=$Discount");
            // Formula that uses the price and the injected discount
            templateSheet.Cells["C1"].Formula = "=A1*B1";

            // Initialize the WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Specify which worksheet holds the variables smart markers
            designer.VariablesWorksheetName = "Variables";

            // Process all smart markers in the workbook
            designer.Process();

            // Recalculate formulas so that the final price reflects the injected discount
            workbook.CalculateFormula();

            // Save the resulting workbook
            string outputPath = "SetVariableDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
