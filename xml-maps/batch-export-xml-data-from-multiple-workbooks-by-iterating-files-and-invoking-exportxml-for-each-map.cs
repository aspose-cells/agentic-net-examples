using System;
using System.IO;
using Aspose.Cells;

namespace BatchXmlExport
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source Excel workbooks
            string inputFolder = @"C:\InputWorkbooks";

            // Folder where the exported XML files will be saved
            string outputFolder = @"C:\ExportedXml";

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Iterate over all Excel files in the input folder
            foreach (string workbookPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load the workbook (uses the Workbook(string) constructor)
                Workbook workbook = new Workbook(workbookPath);

                // Iterate through each XML map defined in the workbook
                foreach (XmlMap xmlMap in workbook.Worksheets.XmlMaps)
                {
                    // Build a unique output file name using workbook name and map name
                    string outputFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_{xmlMap.Name}.xml";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Export the XML data for the current map to the specified file
                    workbook.ExportXml(xmlMap.Name, outputPath);
                }
            }

            Console.WriteLine("Batch XML export completed.");
        }
    }
}