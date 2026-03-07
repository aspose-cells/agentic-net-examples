using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XltmToJsonConverter
    {
        public static void Run()
        {
            // Path to the source XLTM workbook (macro‑enabled template)
            string sourcePath = "template.xltm";

            // Path where the resulting JSON will be saved
            string jsonPath = "output.json";

            // Load the XLTM workbook.
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options (using default settings).
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON.
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"XLTM workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XltmToJsonConverter.Run();
        }
    }
}