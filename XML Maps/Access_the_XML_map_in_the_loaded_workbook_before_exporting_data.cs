using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlFromLoadedWorkbook
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xml";

            Workbook workbook = new Workbook(inputPath);

            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
                workbook.ExportXml(xmlMap.Name, outputPath);
                Console.WriteLine($"XML exported successfully to '{outputPath}'.");
            }
            else
            {
                Console.WriteLine("No XmlMap found in the workbook.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportXmlFromLoadedWorkbook.Run();
        }
    }
}