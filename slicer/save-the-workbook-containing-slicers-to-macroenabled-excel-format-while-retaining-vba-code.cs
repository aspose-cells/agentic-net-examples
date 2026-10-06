// Title: Save a workbook that contains slicers and VBA macros to a macro‑enabled .xlsm file using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx workbook with slicers and a VBA project, then save it as a macro‑enabled .xlsm file while keeping the macros intact using Aspose.Cells in C#. | Convert a slicer‑enabled workbook to a macro‑enabled Excel file without losing any VBA code with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells preserve VBA macros when saving workbook to xlsm C# | C# save workbook with slicers to macro‑enabled Excel using Aspose.Cells | How to keep slicers and macros after converting xlsx to xlsm with Aspose.Cells | Load xlsx containing VBA and slicers and export as xlsm in .NET
// Tags: save workbook as macro-enabled xlsm Aspose.Cells | load xlsx with slicers Aspose.Cells C# | retain VBA project during xlsm export Aspose.Cells | export slicer-enabled workbook to xlsm without losing macros | macro-enabled Excel conversion Aspose.Cells .NET

using Aspose.Cells;

// The example loads an existing .xlsx workbook that includes slicers and a VBA project, then saves it as a macro‑enabled .xlsm file, ensuring that both the slicers and the VBA code are preserved.
class Program
{
    static void Main()
    {
        // Load the existing workbook that contains slicers and VBA code.
        // LoadOptions ensures that macros are preserved (default behavior).
        var loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook workbook = new Workbook("InputWorkbook.xlsx", loadOptions);

        // (Optional) Verify that the workbook contains a VBA project.
        // This step is not required for saving, but demonstrates that macros are present.
        if (workbook.VbaProject != null)
        {
            // VBA project is present and will be retained on save.
        }

        // Save the workbook to a macro‑enabled Excel format (.xlsm) while keeping the VBA code.
        workbook.Save("OutputWorkbook.xlsm", SaveFormat.Xlsm);
    }
}
