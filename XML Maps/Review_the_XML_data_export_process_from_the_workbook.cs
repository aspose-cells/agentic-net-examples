using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlExportReview
    {
        public static void Run()
        {
            // Path to the source Excel workbook
            string excelPath = "Sample.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Check if any XML maps are defined in the workbook
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Use the first XML map for the export demonstration
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
            Console.WriteLine($"Exporting XML using map: {xmlMap.Name}");

            // Export XML to a file
            string outputFilePath = "ExportedData.xml";
            workbook.ExportXml(xmlMap.Name, outputFilePath);
            Console.WriteLine($"XML successfully exported to file: {outputFilePath}");

            // Export XML to a stream (e.g., FileStream)
            string streamOutputPath = "ExportedDataStream.xml";
            using (FileStream fs = new FileStream(streamOutputPath, FileMode.Create, FileAccess.Write))
            {
                workbook.ExportXml(xmlMap.Name, fs);
            }
            Console.WriteLine($"XML successfully exported to stream file: {streamOutputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XmlExportReview.Run();
        }
    }
}