// Title: Get the root element name of the first XML map in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# get root element name of first XML map using Aspose.Cells | Aspose.Cells how to read XmlMap.RootElementName from an .xlsx file | example code for accessing XmlMaps collection in an Aspose.Cells workbook | check if a workbook contains XML maps with Aspose.Cells C# | handle RuntimeBinderException when XmlMaps property is unavailable in Aspose.Cells
// Tags: Aspose.Cells XmlMap RootElementName C# | first XML map extraction Aspose.Cells | XmlMaps collection access Aspose.Cells | Excel workbook XML map root element | runtime binder exception handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, dynamically accesses the XmlMaps collection, ensures at least one XML map is present, and then reads the RootElementName of the first map. It outputs the root element name and includes handling for missing XmlMaps support in older Aspose.Cells versions and other runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Use dynamic to access XmlMaps (may not be available in older versions)
            try
            {
                dynamic dynWorkbook = workbook;
                dynamic xmlMaps = dynWorkbook.XmlMaps;

                if (xmlMaps != null && xmlMaps.Count > 0)
                {
                    // Retrieve the root element name of the first XML map
                    string rootElementName = xmlMaps[0].RootElementName;
                    Console.WriteLine("Root element name: " + rootElementName);
                }
                else
                {
                    Console.WriteLine("No XML maps found in the workbook.");
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // XmlMaps property is not available in this version of Aspose.Cells
                Console.WriteLine("The loaded Aspose.Cells version does not support XML maps.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
