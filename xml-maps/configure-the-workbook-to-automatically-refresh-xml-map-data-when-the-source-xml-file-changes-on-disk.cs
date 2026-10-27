// Title: Enable automatic XML map refresh in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set XmlMap AutoRefresh to true | How to make Excel XML map refresh automatically when source XML changes using Aspose.Cells | Programmatically enable XML map auto‑refresh in a .xlsx file with Aspose.Cells .NET | Aspose.Cells XML map AutoRefresh property example in C#
// Tags: Aspose.Cells XmlMap AutoRefresh C# | automatic XML map refresh Excel .NET | set XmlMap AutoRefresh programmatically | Excel workbook XML map refresh on file change | reflection access XmlMapCollection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The code loads an existing workbook, accesses its XmlMap collection via reflection, sets the first map's AutoRefresh flag to true, and saves the file, allowing the XML map data to update automatically when the source XML file changes.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to access XML maps via reflection (available in newer Aspose.Cells versions)
            var xmlMapsProp = typeof(Workbook).GetProperty("XmlMaps");
            if (xmlMapsProp != null)
            {
                var xmlMaps = xmlMapsProp.GetValue(workbook);
                if (xmlMaps != null)
                {
                    // Use dynamic to interact with the XmlMapCollection without a direct compile‑time reference
                    dynamic maps = xmlMaps;
                    if (maps.Count > 0)
                    {
                        // Enable automatic refresh for the first XML map
                        maps[0].AutoRefresh = true;
                    }
                }
            }
            else
            {
                Console.WriteLine("XML map feature is not available in this Aspose.Cells version.");
            }

            // Save the workbook with the updated setting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
