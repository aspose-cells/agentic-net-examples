// Title: Aspose.Cells for .NET – Load a workbook with a custom LoadFilter to include only needed defined names
// Description: Shows how to configure LoadOptions with a custom LoadFilter that loads defined names solely for selected worksheets (e.g., Sheet1) while loading just the structure for the rest. The sample counts workbook‑ and worksheet‑scoped names, accesses a named range, and saves the workbook, helping to cut load time and memory usage.
// Keywords: Aspose.Cells | .NET | LoadFilter | LoadOptions | defined names | named ranges | selective loading | Excel performance | Workbook scoped names | Worksheet scoped names
// Common Searches: Aspose.Cells load only specific named ranges | custom LoadFilter example .NET | how to load workbook structure without named ranges | filter defined names when opening Excel with Aspose | improve Excel load performance Aspose.Cells
// Developer Intent: Open an Excel file while loading only the required named ranges for calculations.
// Use Cases: Reduce memory consumption by loading defined names only for sheets that need them | Count and enumerate workbook‑scoped and worksheet‑scoped names after selective loading | Retrieve the address of a loaded named range for further processing | Save a workbook after trimming unnecessary name definitions
// AI Prompts: Generate code to extend RequiredNamesLoadFilter to handle multiple worksheets with different name loading strategies. | Show how to use LoadOptions.FilterDefinedNames property to specify an explicit list of named ranges to load. | Explain error handling when a required named range is absent while using a custom LoadFilter in Aspose.Cells. | Provide guidance on measuring performance gains from selective named‑range loading.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to configure LoadOptions with a custom LoadFilter that loads defined names solely for selected worksheets (e.g., Sheet1) while loading just the structure for the rest. The sample counts workbook‑ and worksheet‑scoped names, accesses a named range, and saves the workbook, helping to cut load time and memory usage.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook (template) that contains many defined names
            string sourcePath = "TemplateWithNames.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
                return;
            }

            // Create LoadOptions and assign a custom LoadFilter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new RequiredNamesLoadFilter()
            };

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Retrieve all workbook‑scoped defined names that were loaded
            Name[] workbookScopeNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);
            Console.WriteLine($"Workbook‑scoped names loaded: {workbookScopeNames.Length}");

            // Retrieve worksheet‑scoped names for the first sheet (index 0)
            Name[] sheet1Names = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, 0);
            Console.WriteLine($"Sheet1 names loaded: {sheet1Names.Length}");

            // Demonstrate using a loaded name (if any) to obtain its range
            if (sheet1Names.Length > 0)
            {
                // Resolve ambiguity between Aspose.Cells.Range and System.Range
                Aspose.Cells.Range range = sheet1Names[0].GetRange();
                Console.WriteLine($"First name range address: {range.Address}");
            }

            // Save the workbook after processing
            string outputPath = "ProcessedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Custom LoadFilter that loads only defined names for required sheets
    class RequiredNamesLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only defined names for "Sheet1"; other sheets load structure only
            if (sheet.Name == "Sheet1")
            {
                LoadDataFilterOptions = LoadDataFilterOptions.DefinedNames;
            }
            else
            {
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }
}
