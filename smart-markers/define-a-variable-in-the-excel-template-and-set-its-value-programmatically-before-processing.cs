// Title: Set a smart‑marker variable in an Excel template using Aspose.Cells WorkbookDesigner (C#)
// Common Searches: C# Aspose.Cells set value for a smart marker variable before processing | WorkbookDesigner SetDataSource example for &=MyVar in an .xlsx template | programmatically assign a variable to a smart marker in an Excel file using Aspose.Cells | create Excel template with a smart marker variable and replace it with C# code | Aspose.Cells replace smart marker variable with custom string in .NET
// Tags: Aspose.Cells WorkbookDesigner SetDataSource | smart marker variable binding C# | replace smart marker &=MyVar Aspose.Cells | create Excel template with smart marker | process smart markers after variable assignment

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVariableDemo
{
    // The example creates an Excel template containing the smart‑marker variable '&=MyVar', saves it, loads it with WorkbookDesigner, assigns a value to the variable using SetDataSource, processes the smart markers to replace the marker with the provided value, saves the resulting workbook, and verifies the output by reading the cell value.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the template and the resulting workbook
            string templatePath = "template.xlsx";
            string resultPath = "output.xlsx";

            try
            {
                // ------------------------------------------------------------
                // Step 1: Create an Excel template that contains a variable.
                // The variable is defined as a smart marker: &=MyVar
                // ------------------------------------------------------------
                Workbook templateWorkbook = new Workbook();
                Worksheet templateSheet = templateWorkbook.Worksheets[0];
                Cells templateCells = templateSheet.Cells;

                // Place the smart marker in cell A1
                templateCells["A1"].PutValue("&=MyVar");

                // Optionally add a label next to it
                templateCells["B1"].PutValue("This value is set programmatically.");

                // Save the template to disk
                templateWorkbook.Save(templatePath, SaveFormat.Xlsx);
                Console.WriteLine($"Template workbook created at '{templatePath}'.");

                // ------------------------------------------------------------
                // Step 2: Load the template with WorkbookDesigner, set the variable,
                // and process the smart markers.
                // ------------------------------------------------------------
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                WorkbookDesigner designer = new WorkbookDesigner
                {
                    // Load the template workbook
                    Workbook = new Workbook(templatePath)
                };

                // Bind the variable defined in the template.
                // Use SetDataSource(string variable, object data) for simple variables.
                designer.SetDataSource("MyVar", "Hello, Aspose.Cells!");

                // Process the smart markers – this will replace the marker with the value
                designer.Process();

                // ------------------------------------------------------------
                // Step 3: Save the processed workbook
                // ------------------------------------------------------------
                designer.Workbook.Save(resultPath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed workbook saved at '{resultPath}'.");

                // ------------------------------------------------------------
                // Step 4: Verify the result by reading back the cell value
                // ------------------------------------------------------------
                if (!File.Exists(resultPath))
                    throw new FileNotFoundException($"Result file not found: {resultPath}");

                Workbook resultWorkbook = new Workbook(resultPath);
                string resultValue = resultWorkbook.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine($"Cell A1 after processing contains: \"{resultValue}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
