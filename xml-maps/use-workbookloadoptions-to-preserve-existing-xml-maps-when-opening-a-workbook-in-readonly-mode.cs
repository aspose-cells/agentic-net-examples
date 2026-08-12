// Title: C# – Preserve XML Maps When Loading a Workbook Read‑Only with Aspose.Cells LoadOptions
// Description: Shows how to create a LoadOptions object with a custom LoadFilter that sets LoadDataFilterOptions to Structure | XmlMap, open an .xlsx file in read‑only mode while keeping its XML maps, verify the maps, and save the workbook.
// Keywords: Aspose.Cells | LoadOptions | XML maps | read‑only workbook | LoadFilter | LoadDataFilterOptions | C# | .NET | preserve XML maps | Excel XML mapping | Workbook structure
// Common Searches: Aspose.Cells keep XML maps when opening workbook | LoadOptions read‑only Excel with XML maps C# | custom LoadFilter for XML maps Aspose.Cells | how to preserve XML maps in Aspose.Cells | LoadDataFilterOptions.Structure XmlMap example
// Developer Intent: Open an Excel workbook in read‑only mode while ensuring that any embedded XML maps remain loaded and usable.
// Use Cases: Read‑only processing of an .xlsx file without losing its XML map definitions. | Validating or inspecting XML maps before performing data‑only operations. | Improving performance by loading only the workbook structure and XML maps when the sheet data is not required.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions with a custom LoadFilter to load only the workbook structure and XML maps in read‑only mode. | Explain the effect of LoadDataFilterOptions.Structure | LoadDataFilterOptions.XmlMap on workbook loading and how to confirm that XML maps are loaded. | Provide a step‑by‑step tutorial for preserving XML maps when saving a workbook after read‑only processing with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a LoadOptions object with a custom LoadFilter that sets LoadDataFilterOptions to Structure | XmlMap, open an .xlsx file in read‑only mode while keeping its XML maps, verify the maps, and save the workbook.
class PreserveXmlMapsDemo
{
    static void Main()
    {
        // Create load options and set a custom filter that loads the workbook structure
        // together with the XML maps. This ensures the maps are preserved when the file
        // is opened in read‑only mode.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new XmlMapLoadFilter();

        // Open the workbook with the specified load options.
        // The constructor (string, LoadOptions) follows the provided lifecycle rule.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Verify that XML maps have been loaded.
        Console.WriteLine("Number of XML maps loaded: " + workbook.Worksheets.XmlMaps.Count);
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            XmlMap firstMap = workbook.Worksheets.XmlMaps[0];
            Console.WriteLine("First XML map name: " + firstMap.Name);
        }

        // Save the workbook to a new file, preserving the XML maps.
        // The Save(string) method follows the provided lifecycle rule.
        workbook.Save("output_preserved.xlsx");
    }

    // Custom LoadFilter that tells Aspose.Cells to load the workbook structure
    // and the XML maps. No per‑sheet filtering is required.
    private class XmlMapLoadFilter : LoadFilter
    {
        public XmlMapLoadFilter()
        {
            // Load both the workbook structure and the XML maps.
            this.LoadDataFilterOptions = LoadDataFilterOptions.Structure | LoadDataFilterOptions.XmlMap;
        }

        // Required override; no additional logic needed for each sheet.
        public override void StartSheet(Worksheet sheet) { }
    }
}
