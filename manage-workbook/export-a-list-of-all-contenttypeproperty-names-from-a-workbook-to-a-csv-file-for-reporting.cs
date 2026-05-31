using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class ExportContentTypePropertyNamesToCsv
    {
        public static void Run(string workbookPath, string csvOutputPath)
        {
            try
            {
                // Verify input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(workbookPath);

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(csvOutputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                // Write each ContentTypeProperty name to the CSV file
                using (StreamWriter writer = new StreamWriter(csvOutputPath))
                {
                    foreach (ContentTypeProperty property in workbook.ContentTypeProperties)
                    {
                        writer.WriteLine(property.Name);
                    }
                }

                Console.WriteLine($"Content type property names exported to: {csvOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output CSV path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <workbookPath> <csvOutputPath>");
                return;
            }

            string workbookPath = args[0];
            string csvOutputPath = args[1];

            ExportContentTypePropertyNamesToCsv.Run(workbookPath, csvOutputPath);
        }
    }
}