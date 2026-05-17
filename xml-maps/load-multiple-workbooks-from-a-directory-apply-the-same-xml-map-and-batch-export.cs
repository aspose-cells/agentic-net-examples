using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchExport
{
    public class XmlMapBatchExporter
    {
        /// <summary>
        /// Loads all Excel workbooks from the specified folder, ensures each workbook has the given XML map,
        /// and exports the XML data for each workbook to the output folder.
        /// </summary>
        /// <param name="inputFolder">Folder containing the source Excel files.</param>
        /// <param name="xmlSchemaPath">Path to the XML schema (XSD) that defines the XML map.</param>
        /// <param name="outputFolder">Folder where the exported XML files will be saved.</param>
        public static void Run(string inputFolder, string xmlSchemaPath, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!File.Exists(xmlSchemaPath))
                throw new FileNotFoundException($"XML schema file not found: {xmlSchemaPath}");

            Directory.CreateDirectory(outputFolder);

            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string excelPath in excelFiles)
            {
                Workbook workbook = new Workbook(excelPath);

                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchemaPath);
                    // Optionally rename:
                    // workbook.Worksheets.XmlMaps[mapIndex].Name = "MySharedMap";
                }

                string mapName = workbook.Worksheets.XmlMaps[0].Name;

                string outputXmlPath = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(excelPath) + ".xml");

                workbook.ExportXml(mapName, outputXmlPath);
                workbook.Dispose();
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: <inputFolder> <xmlSchemaPath> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string xmlSchemaPath = args[1];
            string outputFolder = args[2];

            XmlMapBatchExporter.Run(inputFolder, xmlSchemaPath, outputFolder);
        }
    }
}