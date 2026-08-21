// Title: C# – Load Only VBA Project from an Excel Workbook with Aspose.Cells (LoadDataFilterOptions.VBA)
// Description: Demonstrates how to create a LoadOptions object, set its LoadFilter to LoadDataFilterOptions.VBA, and open a macro‑enabled .xlsm file so that only the VBA project is read. The example accesses workbook.VbaProject, prints the project name and module count, and shows how to save the workbook if needed.
// Keywords: Aspose.Cells C# load VBA only | LoadDataFilterOptions.VBA example | read VBA project without worksheets | macro‑enabled workbook loading | Aspose.Cells LoadOptions VBA filter | GitHub Aspose.Cells VBA sample
// Common Searches: Aspose.Cells load only VBA project C# | LoadDataFilterOptions.VBA usage | How to read VBA modules with Aspose.Cells | Skip worksheets when loading .xlsm with Aspose | C# code to extract VBA project from Excel
// Developer Intent: Open an Excel file and retrieve just the VBA project data, avoiding the overhead of loading worksheets or other workbook content.
// Use Cases: Quickly verify whether an uploaded .xlsm contains macros before further processing. | Extract module names and counts for analytics without loading the full workbook. | Modify or copy a VBA project after loading only the VBA data, then save to a new macro‑enabled file.
// AI Prompts: Write C# code that loads only the VBA project from an .xlsm file using Aspose.Cells and prints the source of each module. | Show how to edit a VBA module after loading only VBA data with LoadDataFilterOptions.VBA and save the changes to a new workbook. | Explain how to combine LoadDataFilterOptions.VBA with other filter flags to load VBA plus selected worksheets in a single operation.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to create a LoadOptions object, set its LoadFilter to LoadDataFilterOptions.VBA, and open a macro‑enabled .xlsm file so that only the VBA project is read. The example accesses workbook.VbaProject, prints the project name and module count, and shows how to save the workbook if needed.
class LoadVbaOnlyDemo
{
    static void Main()
    {
        // Path to the macro‑enabled workbook that contains VBA code
        string inputPath = "MacroWorkbook.xlsm";

        // Create LoadOptions and assign a LoadFilter that loads only VBA projects
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.VBA);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the VBA project – it will be loaded, other data may be absent
        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            Console.WriteLine("VBA project loaded successfully.");
            Console.WriteLine($"Project name: {vbaProject.Name}");
            Console.WriteLine($"Number of modules: {vbaProject.Modules.Count}");
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }

        // (Optional) Save the workbook if you need to persist any changes
        // workbook.Save("LoadedVbaOnly.xlsm", SaveFormat.Xlsm);
    }
}
