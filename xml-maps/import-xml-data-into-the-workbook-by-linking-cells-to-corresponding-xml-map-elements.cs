// Title: C# example: Load XML data into an Excel workbook and export to XLSX with Aspose.Cells
// Common Searches: how to load an XML file into an Aspose.Cells workbook in C# | Aspose.Cells C# convert XML document to XLSX file | sample code for importing XML data into Excel using Aspose.Cells for .NET | C# program to read XML and save as Excel workbook with Aspose.Cells | using a temporary XML file with Aspose.Cells to generate an XLSX workbook
// Tags: load XML into Aspose.Cells workbook C# | export workbook to XLSX Aspose.Cells | temporary XML file handling Aspose.Cells | rename worksheet after XML import Aspose.Cells | Aspose.Cells XML map import example

using System;
using System.IO;
using Aspose.Cells;

// The sample writes XML content to a temporary file, loads that file directly into an Aspose.Cells Workbook, optionally renames the first worksheet, saves the workbook as an XLSX file, and then cleans up the temporary XML file.
class Program
{
    static void Main()
    {
        try
        {
            // Sample XML data that matches the schema
            string xmlData = @"<?xml version='1.0'?>
<Root>
  <Item>
    <Name>Alpha</Name>
    <Value>10</Value>
  </Item>
  <Item>
    <Name>Beta</Name>
    <Value>20</Value>
  </Item>
</Root>";

            // Save XML data to a temporary file (Workbook can load XML directly)
            string tempXmlPath = Path.Combine(Path.GetTempPath(), "tempData.xml");
            File.WriteAllText(tempXmlPath, xmlData);

            // Ensure the temporary XML file exists before loading
            if (!File.Exists(tempXmlPath))
                throw new FileNotFoundException("Temporary XML file was not created.", tempXmlPath);

            // Load the XML file into a workbook
            Workbook workbook = new Workbook(tempXmlPath);

            // Optionally, rename the first worksheet
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets[0].Name = "Data";

            // Save the workbook to an Excel file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary XML file if it still exists
            string tempXmlPath = Path.Combine(Path.GetTempPath(), "tempData.xml");
            if (File.Exists(tempXmlPath))
            {
                try
                {
                    File.Delete(tempXmlPath);
                }
                catch
                {
                    // Ignored – best effort cleanup
                }
            }
        }
    }
}
