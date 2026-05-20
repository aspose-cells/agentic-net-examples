using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ImportXmlMapDemo
    {
        public static void Run()
        {
            try
            {
                // Verify that the source XML file exists to avoid FileNotFoundException
                const string xmlPath = "data.xml";
                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine($"Error: XML file '{xmlPath}' not found.");
                    return;
                }

                // Create a new workbook instance
                Workbook wb = new Workbook();

                // Import XML data into the first worksheet ("Sheet1") starting at cell A1 (row 0, column 0)
                wb.ImportXml(xmlPath, "Sheet1", 0, 0);

                // Save the workbook with the imported XML data
                const string outputPath = "ImportedXmlMap.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required for console applications
    public class Program
    {
        public static void Main(string[] args)
        {
            ImportXmlMapDemo.Run();
        }
    }
}