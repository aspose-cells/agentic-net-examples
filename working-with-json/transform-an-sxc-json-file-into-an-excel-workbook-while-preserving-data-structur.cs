using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class SxcJsonToExcel
{
    static void Main()
    {
        // Path to the source SXC JSON file
        string jsonPath = "input.sxc.json";

        // Path for the output Excel workbook
        string excelPath = "output.xlsx";

        // Read the JSON content from the file (use a fallback JSON if the file is missing)
        string jsonContent;
        if (File.Exists(jsonPath))
        {
            jsonContent = File.ReadAllText(jsonPath);
        }
        else
        {
            // Minimal example SXC JSON structure
            jsonContent = @"{
                ""sheets"": [
                    {
                        ""name"": ""Sheet1"",
                        ""rows"": [
                            { ""cells"": [ { ""value"": ""Header1"" }, { ""value"": ""Header2"" } ] },
                            { ""cells"": [ { ""value"": 123 }, { ""value"": 456 } ] }
                        ]
                    }
                ]
            }";
            Console.WriteLine($"Warning: '{jsonPath}' not found. Using fallback JSON data.");
        }

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Import JSON data into the first worksheet starting at cell A1
        JsonUtility.ImportData(jsonContent, workbook.Worksheets[0].Cells, 0, 0, new JsonLayoutOptions());

        // Save the workbook to the specified Excel file
        workbook.Save(excelPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed. Excel file saved to '{excelPath}'.");
    }
}