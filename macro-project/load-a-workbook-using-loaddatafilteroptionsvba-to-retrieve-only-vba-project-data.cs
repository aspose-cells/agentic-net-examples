using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class LoadVbaOnly
{
    static void Main()
    {
        // Path to the source workbook (must be a macro‑enabled file to contain VBA)
        string sourcePath = "input.xlsm";

        // Create LoadOptions and assign a LoadFilter that loads only VBA projects
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.VBA);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Access the VBA project (may be null if the workbook has no VBA)
        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            // Example: output the number of VBA modules present
            Console.WriteLine("VBA Modules count: " + vbaProject.Modules.Count);
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }

        // Save the workbook (preserving the VBA project) to verify loading succeeded
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}