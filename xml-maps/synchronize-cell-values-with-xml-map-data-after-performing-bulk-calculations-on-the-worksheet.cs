// Title: Export Excel cell values to XML maps after bulk formula recalculation using Aspose.Cells for .NET
// AI Prompts: Recalculate all formulas in a workbook, then iterate through the workbook's XmlMaps collection and call ExportData on each map with Aspose.Cells in C#. | Use reflection to obtain the XmlMaps property of a Workbook object and invoke the ExportData method for each XML map to synchronize data after bulk updates. | Load an .xlsx file, perform bulk data changes, recalculate formulas, export the updated cell values to XML maps, and save the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells export updated cell values to XML map after CalculateFormula | C# how to sync Excel worksheet changes with XML maps using Aspose.Cells | Using reflection to access XmlMaps collection in Aspose.Cells workbook | Bulk recalculate formulas and update XML maps in .NET Excel file
// Tags: export data to xml maps Aspose.Cells | recalculate formulas then sync xml map | reflection access XmlMaps C# | bulk formula calculation Excel Aspose.Cells | synchronize worksheet values with xml map .NET

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;

// The example loads an Excel workbook, recalculates all formulas, uses reflection to iterate over each XML map and invoke ExportData, ensuring the XML data stays in sync with the updated cell values, and then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the XML map
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas after bulk data changes
            workbook.CalculateFormula();

            // Export the current cell values back to each XML map to keep the XML data in sync
            // Use reflection to stay compatible with versions that may not expose XmlMaps directly
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp != null)
            {
                var xmlMaps = xmlMapsProp.GetValue(workbook) as IEnumerable;
                if (xmlMaps != null)
                {
                    foreach (var mapObj in xmlMaps)
                    {
                        var exportMethod = mapObj.GetType().GetMethod("ExportData");
                        exportMethod?.Invoke(mapObj, null);
                    }
                }
            }
            else
            {
                Console.WriteLine("The loaded workbook does not contain any XML maps or the XmlMaps property is unavailable.");
            }

            // Save the workbook with the synchronized XML map data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
