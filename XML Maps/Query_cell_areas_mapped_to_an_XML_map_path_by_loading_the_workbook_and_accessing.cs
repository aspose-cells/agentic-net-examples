using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapQueryDemo
    {
        public static void Run(string inputPath, string outputPath, string xmlMapPath)
        {
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps are defined in the workbook.");
                return;
            }

            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
            ArrayList cellAreas = worksheet.XmlMapQuery(xmlMapPath, xmlMap);

            if (cellAreas.Count > 0)
            {
                Console.WriteLine($"Found {cellAreas.Count} cell area(s) linked to path \"{xmlMapPath}\":");
                foreach (CellArea area in cellAreas)
                {
                    Console.WriteLine($"StartRow: {area.StartRow}, StartColumn: {area.StartColumn}, " +
                                      $"EndRow: {area.EndRow}, EndColumn: {area.EndColumn}");

                    Cell topLeftCell = worksheet.Cells[area.StartRow, area.StartColumn];
                    Console.WriteLine($"Top‑left cell value: {topLeftCell.StringValue}");
                }
            }
            else
            {
                Console.WriteLine($"No cell areas are linked to the path \"{xmlMapPath}\".");
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <inputPath> <outputPath> <xmlMapPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string xmlMapPath = args[2];

            XmlMapQueryDemo.Run(inputPath, outputPath, xmlMapPath);
        }
    }
}