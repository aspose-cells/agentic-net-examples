// Title: Filter and Remove Underscore‑Prefixed Defined Names When Loading a Workbook – Aspose.Cells for .NET
// Description: This example creates a workbook with visible and '_'‑prefixed defined names, saves it, then reloads the file using a LoadFilter configured for LoadDataFilterOptions.DefinedNames. It programmatically removes all names that start with '_' via NameCollection.Remove, verifies their absence, and saves the cleaned workbook.
// Keywords: Aspose.Cells | LoadFilter | DefinedNames | C# | remove defined names | underscore prefix | named ranges | LoadDataFilterOptions | Workbook loading | verify removal
// Common Searches: Aspose.Cells load only defined names | remove underscore named ranges Aspose.Cells | filter defined names during workbook load .NET | check for hidden defined names after deletion | LoadOptions LoadFilter example
// Developer Intent: Load a workbook, delete every defined name that begins with '_' and confirm that none remain.
// Use Cases: Create a workbook containing both visible and '_'‑prefixed defined names and persist it. | Reload the workbook with LoadOptions.LoadFilter set to DefinedNames to limit the data read. | Identify and delete all defined names whose Text starts with '_' using NameCollection.Remove. | Validate that no underscore‑prefixed names exist before saving the cleaned file. | Reuse the cleaned workbook for further processing or distribution.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells using LoadFilter to load only defined names, then removes any names beginning with '_' and confirms their removal. | Explain the purpose of LoadDataFilterOptions.DefinedNames in Aspose.Cells and how it improves performance when only the name collection is needed. | Provide alternative approaches to exclude underscore‑prefixed defined names without loading the entire name collection, such as custom LoadFilter logic or post‑load filtering. | Generate a PowerShell script that uses Aspose.Cells for .NET to batch‑process multiple workbooks, removing underscore‑prefixed defined names.

using System;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsDefinedNameFilter
{
    // This example creates a workbook with visible and '_'‑prefixed defined names, saves it, then reloads the file using a LoadFilter configured for LoadDataFilterOptions.DefinedNames. It programmatically removes all names that start with '_' via NameCollection.Remove, verifies their absence, and saves the cleaned workbook.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and add defined names (some start with "_")
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Add visible name
            int idxVisible = wb.Worksheets.Names.Add("VisibleRange");
            Name visibleName = wb.Worksheets.Names[idxVisible];
            visibleName.RefersTo = "=Sheet1!$A$1";

            // Add hidden names (start with "_")
            int idxHidden1 = wb.Worksheets.Names.Add("_HiddenRange1");
            Name hiddenName1 = wb.Worksheets.Names[idxHidden1];
            hiddenName1.RefersTo = "=Sheet1!$B$1";

            int idxHidden2 = wb.Worksheets.Names.Add("_HiddenRange2");
            Name hiddenName2 = wb.Worksheets.Names[idxHidden2];
            hiddenName2.RefersTo = "=Sheet1!$C$1";

            // Save the workbook to a temporary file
            string filePath = "DefinedNamesDemo.xlsx";
            wb.Save(filePath);
            wb.Dispose();

            // ---------------------------------------------------------------
            // 2. Load the workbook with a LoadFilter that loads defined names
            // ---------------------------------------------------------------
            // Create a LoadFilter that enables loading of defined names only
            LoadFilter loadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

            // Assign the filter to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = loadFilter;

            // Load the workbook using the options
            Workbook loadedWb = new Workbook(filePath, loadOptions);

            // ---------------------------------------------------------------
            // 3. Remove all defined names that start with "_"
            // ---------------------------------------------------------------
            NameCollection names = loadedWb.Worksheets.Names;

            // Find names beginning with '_' and collect their texts
            string[] namesToRemove = names
                .Cast<Name>()
                .Where(n => n.Text.StartsWith("_"))
                .Select(n => n.Text)
                .ToArray();

            // Use the provided Remove(string[]) method to delete them
            if (namesToRemove.Length > 0)
            {
                names.Remove(namesToRemove);
            }

            // ---------------------------------------------------------------
            // 4. Verify that no names starting with '_' remain
            // ---------------------------------------------------------------
            bool anyUnderscoreNames = names
                .Cast<Name>()
                .Any(n => n.Text.StartsWith("_"));

            Console.WriteLine("Names starting with '_' present after removal: " + anyUnderscoreNames);
            Console.WriteLine("Remaining defined names:");
            foreach (Name n in names)
            {
                Console.WriteLine("- " + n.Text);
            }

            // Optionally save the cleaned workbook
            loadedWb.Save("DefinedNamesDemo_Cleaned.xlsx");
            loadedWb.Dispose();
        }
    }
}
