// Title: Add a Workbook_BeforePrint VBA event to a macro‑enabled Excel workbook with Aspose.Cells for .NET to block printing
// AI Prompts: Generate C# code that uses Aspose.Cells to load an existing .xlsm file, create a ThisWorkbook class module, insert a Workbook_BeforePrint procedure that sets Cancel = True and shows a message, then save the workbook as a macro‑enabled file. | Demonstrate how to programmatically attach VBA code that disables the Print command by adding a class module through the Aspose.Cells .NET API.
// Common Searches: how to programmatically disable printing in an .xlsm workbook using Aspose.Cells C# | add Workbook_BeforePrint event to existing macro enabled Excel file with .NET | insert VBA code into Excel file with Aspose.Cells to cancel print operation | Aspose.Cells C# add VBA class module to prevent user from printing | save changes to macro enabled workbook after adding VBA event using Aspose.Cells
// Tags: Aspose.Cells add VBA class module C# | Workbook_BeforePrint event Aspose.Cells | disable Excel printing via VBA .NET | save macro enabled workbook Aspose.Cells | programmatic VBA insertion .xlsm C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads a macro‑enabled workbook, verifies the VBA project, adds a ThisWorkbook class module containing a Workbook_BeforePrint procedure that cancels printing and displays a message, and saves the modified file as PrintDisabled.xlsm.
class Program
{
    static void Main()
    {
        try
        {
            // Path to a macro‑enabled template workbook that already contains a VBA project
            const string templatePath = "MacroTemplate.xlsm";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file '{templatePath}' not found. Please provide a macro‑enabled workbook.");
                return;
            }

            // Load the template workbook (must be .xlsm to have a VBA project)
            Workbook workbook = new Workbook(templatePath);

            // Ensure a VBA project exists; if not, inform the user and exit
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("Error: The loaded workbook does not contain a VBA project. Use a macro‑enabled template.");
                return;
            }

            // Add a class module named "ThisWorkbook" to handle workbook‑level events
            // In some Aspose.Cells versions Add returns the module index (int)
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "ThisWorkbook");
            VbaModule thisWorkbookModule = workbook.VbaProject.Modules[moduleIndex];

            // VBA code that cancels any print attempt
            string vbaCode = @"
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Cancel = True
    MsgBox ""Printing is disabled by policy.""
End Sub
";

            // Assign the VBA code to the newly created module
            thisWorkbookModule.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file so the VBA code is retained
            const string outputPath = "PrintDisabled.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
