using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class NumbersJsonToCsv
{
    static void Main()
    {
        string jsonPath = "input.numbers.json";
        string jsonContent;

        if (File.Exists(jsonPath))
        {
            jsonContent = File.ReadAllText(jsonPath);
        }
        else
        {
            // Sample JSON data if the file does not exist
            jsonContent = "[{\"Name\":\"Alice\",\"Score\":85},{\"Name\":\"Bob\",\"Score\":92}]";
        }

        Workbook workbook = new Workbook();

        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        JsonUtility.ImportData(jsonContent, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

        string csvOutputPath = "output.csv";
        workbook.Save(csvOutputPath, SaveFormat.Csv);
    }
}