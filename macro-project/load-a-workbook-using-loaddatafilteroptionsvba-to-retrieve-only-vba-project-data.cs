using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class LoadVbaOnlyDemo
{
    static void Main()
    {
        // Path to the macro‑enabled workbook
        string filePath = "sample.xlsm";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only VBA project data
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.VBA);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Access the VBA project (may be null if none exists)
        VbaProject vbaProject = workbook.VbaProject;

        // Indicate whether the VBA project was loaded
        Console.WriteLine("VBA Project loaded: " + (vbaProject != null));

        // If a VBA project exists, display the number of modules it contains
        if (vbaProject != null)
        {
            Console.WriteLine("Number of VBA modules: " + vbaProject.Modules.Count);
        }
    }
}