using System;
using Aspose.Cells;

class LoadWorkbookWithoutMacro
{
    static void Main()
    {
        // Configure LoadOptions to exclude VBA projects
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter();

        // Load a macro‑enabled workbook with the specified options
        Workbook workbook = new Workbook("sample.xlsm", loadOptions);

        // Verify that macros are not present after loading
        Console.WriteLine("HasMacro after load: " + workbook.HasMacro);

        // Save the workbook as a macro‑free file (optional)
        workbook.Save("sample_no_macro.xlsx", SaveFormat.Xlsx);
    }

    // Custom LoadFilter that loads only the workbook structure, omitting VBA
    private class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only the structural information; VBA (macro) data is excluded
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}