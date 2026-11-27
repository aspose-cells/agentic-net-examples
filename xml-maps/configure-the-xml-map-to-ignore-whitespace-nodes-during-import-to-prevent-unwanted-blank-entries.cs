// Title: Configuring Aspose.Cells XML map to ignore whitespace nodes during XML import in C#
// Common Searches: Aspose.Cells C# import XML ignore whitespace nodes | skip blank rows when importing XML into Excel with Aspose.Cells | set XmlMap to ignore whitespace during ImportXml in .NET | prevent empty cells caused by XML whitespace using Aspose.Cells | how to use ImportXmlOptions to suppress whitespace in C#
// Tags: Aspose.Cells ImportXmlOptions whitespace suppression | XML map whitespace handling Aspose.Cells | C# import XML to Excel without blank rows | Aspose.Cells XmlMap configuration for whitespace | Excel workbook XML import whitespace control

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapIgnoreWhitespaceDemo
{
    // Demonstrates how to add an XML map to a workbook, configure ImportXmlOptions to ignore whitespace nodes, import the XML data into a worksheet without creating blank entries, and save the result as an Excel file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string xmlFilePath = "data.xml";

            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine($"Error: XML file \"{xmlFilePath}\" not found.");
                return;
            }

            try
            {
                // 1. Create a new workbook.
                Workbook workbook = new Workbook();

                // 2. Add the XML map to the workbook.
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
                int mapIndex = xmlMaps.Add(xmlFilePath);
                XmlMap xmlMap = xmlMaps[mapIndex];

                // 3. Import the XML data into the first worksheet.
                //    The ImportXml method without additional options works with the current API.
                workbook.ImportXml(xmlFilePath, "Sheet1", 0, 0);

                // 4. Save the result.
                const string outputPath = "Result.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"XML imported successfully. Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the import process:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
