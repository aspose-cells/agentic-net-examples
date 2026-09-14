// Title: Load an XML file into an Aspose.Cells workbook, rename the first worksheet, and save as XLSX using C#
// AI Prompts: Use Aspose.Cells LoadOptions with LoadFormat.Xml to read an XML file into a Workbook, rename the first sheet, and export the result to a .xlsx file in C#. | Write a C# method that verifies an XML data file exists, loads it into a new Workbook via Aspose.Cells, changes the initial worksheet name, and saves the workbook as Excel with proper error handling. | Demonstrate how to convert XML to an Excel workbook while customizing the worksheet title using Aspose.Cells for .NET.
// Common Searches: asp.net load xml into workbook with aspose.cells and rename sheet | c# aspose.cells import xml data and export to xlsx | how to change worksheet name after loading xml using aspose.cells | using LoadOptions LoadFormat.Xml to convert xml to excel in c# | aspose.cells error handling when loading xml file
// Tags: aspose.cells loadoptions xml to workbook | c# rename worksheet after xml import | aspose.cells convert xml to xlsx | error handling loading xml with aspose.cells | aspose.cells workbook save as xlsx

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the presence of a data.xml file, loads its contents into a new Aspose.Cells Workbook using LoadOptions with LoadFormat.Xml, renames the first worksheet to "ImportedData", and saves the workbook as output.xlsx, with comprehensive exception handling for file and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the XML file that contains the data to be imported
            string xmlPath = "data.xml";

            // Verify that the XML file exists to avoid FileNotFoundException
            if (!File.Exists(xmlPath))
            {
                Console.WriteLine($"Error: XML file \"{xmlPath}\" not found.");
                return;
            }

            // Load the XML data into a new workbook using LoadOptions for XML format
            Workbook workbook;
            try
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
                workbook = new Workbook(xmlPath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading XML data: {ex.Message}");
                return;
            }

            // Rename the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ImportedData";

            // Save the workbook with the imported XML data
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
