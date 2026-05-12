using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsvDemo
{
    class Program
    {
        static void Main()
        {
            string json = @"{
                ""Employees"": [
                    { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                    { ""Name"": ""Anna"", ""Age"": 25, ""City"": ""London"" },
                    { ""Name"": ""Mike"", ""Age"": 35, ""City"": ""Paris"" }
                ]
            }";

            Workbook workbook = new Workbook();

            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true
            };

            JsonUtility.ImportData(json, workbook.Worksheets[0].Cells, 0, 0, jsonOptions);

            TxtSaveOptions csvOptions = new TxtSaveOptions
            {
                Separator = ';'
            };

            workbook.Save("Employees.csv", csvOptions);
        }
    }
}