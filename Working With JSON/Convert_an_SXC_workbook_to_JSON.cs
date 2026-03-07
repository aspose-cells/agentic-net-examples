using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    public class SxcToJsonConverter
    {
        public static void Run()
        {
            // Path to the source SXC workbook
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.sxc");

            // Path where the resulting JSON will be saved
            string destinationPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.json");

            // If the source file does not exist, create a sample workbook and save it as SXC
            if (!File.Exists(sourcePath))
            {
                Workbook sampleWb = new Workbook();
                Worksheet sheet = sampleWb.Worksheets[0];
                sheet.Name = "SampleSheet";

                // Add header row
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");

                // Add some data
                sheet.Cells["A2"].PutValue("Data1");
                sheet.Cells["B2"].PutValue("Data2");

                // Save as SXC (OpenOffice Calc) format
                sampleWb.Save(sourcePath, SaveFormat.Sxc);
            }

            // Load the SXC workbook
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options (default options are sufficient)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON using the configured options
            workbook.Save(destinationPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {destinationPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SxcToJsonConverter.Run();
        }
    }
}