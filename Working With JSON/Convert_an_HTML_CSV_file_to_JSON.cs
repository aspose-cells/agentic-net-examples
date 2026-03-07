using System;
using Aspose.Cells;

class HtmlCsvToJson
{
    static void Main()
    {
        string csvPath = "input.csv";
        string jsonPath = "output.json";

        // Load CSV file into a workbook
        Workbook workbook = new Workbook(csvPath, new LoadOptions(LoadFormat.Csv));

        // Save the workbook as JSON
        workbook.Save(jsonPath, SaveFormat.Json);
    }
}