// Title: Remove an unwanted XML map from an Excel workbook by index using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, checks for an XML map at a given index, removes it, and saves the workbook. | Show a version‑safe approach that uses reflection to access the XmlMaps collection and delete a specific XML map without compile‑time dependencies. | Demonstrate how to verify the presence of XmlMaps in a workbook and programmatically remove the first map when it exists.
// Common Searches: Aspose.Cells remove xml map by index C# | delete unwanted XML map from Excel workbook programmatically | how to use reflection to access XmlMaps collection in Aspose.Cells
// Tags: aspocells xmlmap removal using reflection | c# delete xml map from workbook | excel xml map collection manipulation | version‑compatible xmlmaps access | dynamic xmlmap handling aspocells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// // Loads an Excel file, uses reflection to obtain the XmlMaps collection if available, removes the XML map at index 0 when present, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Attempt to remove an XML map if the XmlMaps property is available in this version
            const int mapIndex = 0;
            PropertyInfo xmlMapsProp = typeof(Workbook).GetProperty("XmlMaps");
            if (xmlMapsProp != null)
            {
                // Use dynamic to work with the XmlMapCollection without a direct compile‑time reference
                dynamic xmlMaps = xmlMapsProp.GetValue(workbook);
                if (xmlMaps != null && mapIndex >= 0 && mapIndex < xmlMaps.Count)
                {
                    xmlMaps.RemoveAt(mapIndex);
                }
            }
            else
            {
                Console.WriteLine("XmlMaps property is not supported in the current Aspose.Cells version.");
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
