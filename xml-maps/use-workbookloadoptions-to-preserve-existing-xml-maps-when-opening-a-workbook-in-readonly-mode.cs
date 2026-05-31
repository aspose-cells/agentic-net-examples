using System;
using System.IO;
using Aspose.Cells;

namespace SampleApp
{
    class PreserveXmlMapsReadOnly
    {
        // Custom LoadFilter that tells the loader to include XML maps (and all other data)
        private class XmlMapLoadFilter : LoadFilter
        {
            public override void StartSheet(Worksheet sheet)
            {
                // Load everything for the sheet, ensuring XmlMap data is preserved
                this.LoadDataFilterOptions = LoadDataFilterOptions.All | LoadDataFilterOptions.XmlMap;
            }
        }

        static void Main()
        {
            string filePath = "SampleWithXmlMaps.xlsx";

            // Prevent FileNotFoundException by checking file existence first
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Create LoadOptions and attach the custom filter
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new XmlMapLoadFilter()
                };

                // Open the workbook with the specified options (read‑only – we simply do not modify or save it)
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Verify that XML maps have been loaded
                Console.WriteLine("XML maps loaded: " + workbook.Worksheets.XmlMaps.Count);
                foreach (XmlMap map in workbook.Worksheets.XmlMaps)
                {
                    Console.WriteLine($"Map Name: {map.Name}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}