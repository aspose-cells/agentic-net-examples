// Title: Inject a runtime value into a smart marker variable using Aspose.Cells SetVariable and WorkbookDesigner in C#
// AI Prompts: Generate C# code that creates a variables worksheet, assigns a runtime string to a smart‑marker variable with SetVariable, and processes the workbook using WorkbookDesigner. | Show the step‑by‑step configuration of the VariablesWorksheetName property, placement of a &=$MyVariable placeholder, injection of a value, and saving of the final .xlsx file. | Explain how Aspose.Cells SetVariable can dynamically replace smart marker variables across sheets in a C# application.
// Common Searches: Aspose.Cells C# SetVariable smart marker runtime value example | how to use WorkbookDesigner variables worksheet for smart markers in .NET | replace smart marker variable with dynamic data using Aspose.Cells SetVariable | C# code sample for injecting values into smart marker variables before processing | saving workbook after processing smart markers with SetVariable in Aspose.Cells
// Tags: Aspose.Cells SetVariable smart marker | WorkbookDesigner variables worksheet C# | dynamic smart marker substitution .xlsx | inject runtime value into smart marker | process smart markers with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSetVariableDemo
{
    // Demonstrates creating a workbook, adding a Variables sheet with a &=$MyVariable placeholder, injecting a runtime value into that cell, configuring WorkbookDesigner to use the variables sheet, adding a Template sheet that references the variable smart marker, processing the markers to substitute the value, and saving the result as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet that will hold the variable smart marker
                Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
                // Place a placeholder for the variable (required for SetVariable processing)
                variablesSheet.Cells["A1"].PutValue("&=$MyVariable");
                // Inject the runtime value directly into the variable cell
                variablesSheet.Cells["A1"].PutValue("Injected Runtime Value");

                // Initialize WorkbookDesigner with the workbook and specify the variables sheet
                WorkbookDesigner designer = new WorkbookDesigner(workbook)
                {
                    VariablesWorksheetName = "Variables"
                };

                // Add a template worksheet that uses the variable smart marker
                Worksheet templateSheet = workbook.Worksheets.Add("Template");
                templateSheet.Cells["A1"].PutValue("&=$MyVariable");

                // Process the smart markers – the variable value will be substituted
                designer.Process();

                // Define output file path
                string outputPath = "SetVariableDemo.xlsx";

                // Ensure the output directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully at '{Path.GetFullPath(outputPath)}' with SetVariable applied.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
