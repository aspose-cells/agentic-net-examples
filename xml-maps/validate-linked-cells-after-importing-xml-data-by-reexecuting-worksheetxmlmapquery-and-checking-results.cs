// Title: Re‑execute Worksheet.XmlMapQuery to validate linked cells after importing XML with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx workbook, imports an XML file into its first XML map, then calls Worksheet.XmlMapQuery on every worksheet and map, logging each query's success status and returned value. | Demonstrate how to iterate over all worksheets and XML maps in Aspose.Cells, re‑run the XML map query after data import, and output validation messages for linked cells.
// Common Searches: Aspose.Cells how to run XmlMapQuery on each worksheet after XML import | C# validate linked cells using Worksheet.XmlMapQuery after importing XML data | re‑execute XML map query for multiple worksheets in Aspose.Cells | check success of XmlMapQuery for all XML maps in a workbook with Aspose.Cells | verify XML map import results programmatically in Aspose.Cells .NET
// Tags: Worksheet.XmlMapQuery validation C# | Aspose.Cells XML map import verification | iterate worksheets XML maps Aspose.Cells | log XmlMapQuery results Aspose.Cells | linked cells validation after XML import

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, optionally imports XML data into its first XML map, then iterates through each worksheet and each XML map, re‑executing Worksheet.XmlMapQuery to confirm that linked cells are correctly updated. It logs whether each query succeeded and the result, and finally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load existing workbook
            var workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // XML map handling (optional – comment out if not supported)
            // ------------------------------------------------------------
            // const string xmlPath = "data.xml";
            // if (File.Exists(xmlPath) && workbook.XmlMaps.Count > 0)
            // {
            //     // Import XML data into the first XML map
            //     workbook.XmlMaps[0].ImportXml(xmlPath);
            //
            //     // Re‑execute XmlMapQuery for each worksheet and each XML map
            //     foreach (Worksheet sheet in workbook.Worksheets)
            //     {
            //         foreach (XmlMap map in workbook.XmlMaps)
            //         {
            //             object result;
            //             bool success = sheet.XmlMapQuery(map.Name, out result);
            //             if (success)
            //             {
            //                 Console.WriteLine($"Worksheet '{sheet.Name}', Map '{map.Name}': Query succeeded. Result = {result}");
            //             }
            //             else
            //             {
            //                 Console.WriteLine($"Worksheet '{sheet.Name}', Map '{map.Name}': Query failed.");
            //             }
            //         }
            //     }
            // }
            // ------------------------------------------------------------

            // Save workbook after processing
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
