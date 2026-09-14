// Title: Retrieve the VBA project from an Excel workbook and list its macro modules using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and accesses the workbook.VbaProject property. | Provide a snippet that iterates over VbaProject.Modules and prints each module's name. | Show how to verify that a workbook contains a VBA project before enumerating its modules with Aspose.Cells.
// Common Searches: how to get VBA project from an Excel file using Aspose.Cells C# | list all macro modules in a .xlsx workbook with Aspose.Cells .NET | Aspose.Cells retrieve VbaProject and enumerate modules example | C# read VBA macros from workbook using Aspose.Cells | check for VBA project in workbook before accessing modules Aspose.Cells
// Tags: Aspose.Cells retrieve VbaProject | enumerate VBA modules Aspose.Cells | list macro module names C# | access workbook VBA project .NET | read VBA project from .xlsx programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an Excel file (input.xlsx) into an Aspose.Cells Workbook, obtains its VbaProject, and iterates through the Modules collection, outputting each module's name to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Example usage: list all VBA modules in the project
        foreach (VbaModule module in vbaProject.Modules)
        {
            Console.WriteLine($"Module Name: {module.Name}");
        }
    }
}
