// Title: C# – Extract VBA Project Description from a Macro‑Enabled .xlsm with Aspose.Cells
// Description: Loads a macro‑enabled workbook using Aspose.Cells for .NET, checks for a VBA project, and reads the project's Name property (used as the description). The code prints the description or reports that no VBA project exists.
// Keywords: Aspose.Cells VBA project name | C# read .xlsm macro metadata | extract VBA description Aspose | check workbook.HasMacro | .NET load macro‑enabled Excel
// Common Searches: How to get VBA project name from an xlsm file in C# | Aspose.Cells read VBA project description | Check if Excel workbook contains macros with Aspose | Retrieve macro project metadata using Aspose.Cells .NET
// Developer Intent: Read a macro‑enabled Excel file and obtain the VBA project's description (Name) for logging or UI display.
// Use Cases: Validate uploaded .xlsm files and log their VBA project names before further processing. | Create an inventory of Excel workbooks with associated VBA project descriptions for documentation. | Show the VBA project name in an application UI to inform users about embedded macros.
// AI Prompts: Generate C# code that opens an .xlsm file with Aspose.Cells, verifies the presence of a VBA project, and returns the project's Name property. | Provide an example that extracts the VBA project description from a workbook and gracefully handles the case where no VBA project is present, using Aspose.Cells for .NET. | Show how to combine VBA project name extraction with other workbook metadata (author, creation date) in a single Aspose.Cells routine.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a macro‑enabled workbook using Aspose.Cells for .NET, checks for a VBA project, and reads the project's Name property (used as the description). The code prints the description or reports that no VBA project exists.
class Program
{
    static void Main()
    {
        // Load the macro-enabled workbook from file
        string filePath = "input.xlsm";
        Workbook workbook = new Workbook(filePath);

        // Verify that the workbook contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Extract the VBA project's name (used here as the description)
            string vbaDescription = workbook.VbaProject.Name;
            Console.WriteLine("VBA Project Description: " + vbaDescription);
        }
        else
        {
            Console.WriteLine("The loaded workbook does not contain a VBA project.");
        }
    }
}
