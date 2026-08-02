using System;
using Aspose.Cells;

class PreserveXmlMapsReadOnly
{
    static void Main()
    {
        // Create LoadOptions and assign a custom LoadFilter that includes XmlMap data.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new XmlMapLoadFilter();

        // Open the workbook with the specified LoadOptions.
        // This loads the workbook in a read‑only manner (no modifications are made).
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Verify that XML maps have been preserved.
        Console.WriteLine("Number of XML maps: " + workbook.Worksheets.XmlMaps.Count);
        foreach (XmlMap map in workbook.Worksheets.XmlMaps)
        {
            Console.WriteLine($"Map Name: {map.Name}");
        }

        // Optionally save a copy of the workbook (the original XML maps remain intact).
        workbook.Save("output_copy.xlsx");
    }

    // Custom LoadFilter that ensures XmlMap data is loaded together with the rest of the workbook.
    private class XmlMapLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load all sheet data and explicitly include XML maps.
            LoadDataFilterOptions = LoadDataFilterOptions.All | LoadDataFilterOptions.XmlMap;
        }
    }
}