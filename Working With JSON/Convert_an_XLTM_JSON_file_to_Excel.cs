using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an XLTM template
            string sourceJsonPath = "template.json";

            // Ensure the JSON file exists (create a simple example if it does not)
            if (!File.Exists(sourceJsonPath))
            {
                string sampleJson = @"{
                    ""Sheets"": [
                        {
                            ""Name"": ""Sheet1"",
                            ""Rows"": [
                                {
                                    ""Cells"": [
                                        { ""Column"": 0, ""Value"": ""Hello"" },
                                        { ""Column"": 1, ""Value"": 123 }
                                    ]
                                }
                            ]
                        }
                    ]
                }";
                File.WriteAllText(sourceJsonPath, sampleJson);
            }

            // Desired output Excel file (XLSX format)
            string outputExcelPath = "ConvertedTemplate.xlsx";

            // LoadOptions specifying that the source file is in JSON format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);

            // Load the JSON workbook
            Workbook workbook = new Workbook(sourceJsonPath, loadOptions);

            // Save the workbook as XLSX
            workbook.Save(outputExcelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to: {outputExcelPath}");
        }
    }
}