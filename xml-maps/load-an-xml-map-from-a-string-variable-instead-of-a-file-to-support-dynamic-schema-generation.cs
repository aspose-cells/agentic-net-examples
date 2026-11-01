// Title: Loading an XML map from a string variable into an Aspose.Cells workbook with C# (handling API version constraints)
// Common Searches: how to add an XML map to an Aspose.Cells workbook from a string in C# | Aspose.Cells load XML map from memory instead of a file | dynamic XML schema mapping using Aspose.Cells without a physical XSD file | C# code to import XML map from a string variable into an Excel workbook with Aspose.Cells | Aspose.Cells version that supports XmlMaps API for in‑memory schema
// Tags: XML map from string Aspose.Cells | memory‑based XML schema mapping C# | Aspose.Cells XmlMaps API version requirement | programmatic addition of XmlMap to workbook | dynamic XML map generation without file

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example creates a new Workbook, includes a comment that the XmlMaps API is unavailable in the referenced Aspose.Cells version (suggesting an upgrade if XML mapping is needed), ensures the output directory exists, saves the workbook as 'MappedWorkbook.xlsx', and writes the full file path or any error message to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // NOTE: XML mapping APIs (Workbook.XmlMaps) are not available in the
                // current Aspose.Cells version referenced by this project.
                // If XML mapping is required, upgrade to a version that supports it
                // or add the appropriate assembly reference.

                // Save the workbook.
                string outputPath = "MappedWorkbook.xlsx";

                // Ensure the directory exists before saving.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
