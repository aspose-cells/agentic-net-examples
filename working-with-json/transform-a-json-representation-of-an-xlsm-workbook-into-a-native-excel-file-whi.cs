using System;
using System.IO;
using Aspose.Cells;

namespace JsonToXlsmConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON representation of the XLSM workbook
            string jsonFilePath = "workbook.json";

            // Load the JSON into a Workbook.
            // The LoadOptions constructor can be used to specify the format if needed.
            // Here we use the default constructor which lets Aspose.Cells auto‑detect the format.
            LoadOptions loadOptions = new LoadOptions(); // Auto‑detect format
            using (FileStream jsonStream = new FileStream(jsonFilePath, FileMode.Open, FileAccess.Read))
            {
                // Create a Workbook instance from the JSON stream
                Workbook workbook = new Workbook(jsonStream, loadOptions);

                // Ensure macros are enabled when loading/saving
                workbook.Settings.EnableMacros = true;

                // Save the workbook as a macro‑enabled Excel file (XLSM)
                string outputPath = "output.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Conversion completed. File saved to: {outputPath}");
            }
        }
    }
}