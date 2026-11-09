// Title: List all VBA modules in an Excel workbook and display their names with Aspose.Cells for .NET (C#)
// Common Searches: C# code to get VBA module names from an .xlsx file using Aspose.Cells | how to iterate through VbaProject modules in Aspose.Cells .NET | retrieve list of VBA modules in Excel workbook with Aspose.Cells library | print VBA module names to console using Aspose.Cells for C# | check if workbook contains VBA project and list modules Aspose.Cells
// Tags: enumerate vba modules Aspose.Cells | read vba project from xlsx C# | list vba module names console output | aspocells workbook vba module extraction | c# aspocells vba project handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an existing .xlsx file with Aspose.Cells, checks for a VBA project, and if present iterates through the VbaProject.Modules collection, writing each module's Name to the console; otherwise it reports that no VBA project exists and handles errors gracefully.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the VBA project (if any) from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.Modules != null && vbaProject.Modules.Count > 0)
            {
                // Iterate through each VBA module and output its name
                foreach (VbaModule module in vbaProject.Modules)
                {
                    Console.WriteLine(module.Name);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
