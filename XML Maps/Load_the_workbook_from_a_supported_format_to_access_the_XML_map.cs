using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that contains XML maps
            string inputPath = "input.xlsx";

            // Load the workbook (uses the provided workbook-load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of XML maps in the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Display information about the XML maps
            Console.WriteLine($"Number of XML maps: {xmlMaps.Count}");
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                Console.WriteLine($"Map {i + 1}: Name = {xmlMaps[i].Name}");
            }

            // Save the workbook (uses the provided workbook-save rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}