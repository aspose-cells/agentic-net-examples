using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

public class DataTransformer
{
    public void Transform(string inputPath, string outputPath)
    {
        Workbook workbook = new Workbook();

        string ext = Path.GetExtension(inputPath).ToLowerInvariant();

        if (ext == ".xml")
        {
            workbook.ImportXml(inputPath, "Sheet1", 0, 0);
        }
        else if (ext == ".json")
        {
            string jsonContent = File.ReadAllText(inputPath);
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
                throw new InvalidOperationException("Root element must be a JSON array.");

            List<string> headers = new List<string>();
            foreach (JsonProperty prop in root[0].EnumerateObject())
                headers.Add(prop.Name);

            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells.ImportObjectArray(headers.ToArray(), 0, 0, false);

            int currentRow = 1;
            foreach (JsonElement element in root.EnumerateArray())
            {
                List<object?> rowValues = new List<object?>();
                foreach (string header in headers)
                {
                    if (element.TryGetProperty(header, out JsonElement value))
                    {
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.Number:
                                if (value.TryGetInt64(out long l))
                                    rowValues.Add(l);
                                else if (value.TryGetDouble(out double d))
                                    rowValues.Add(d);
                                else
                                    rowValues.Add(value.GetRawText());
                                break;
                            case JsonValueKind.String:
                                rowValues.Add(value.GetString());
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                rowValues.Add(value.GetBoolean());
                                break;
                            case JsonValueKind.Null:
                                rowValues.Add(null);
                                break;
                            default:
                                rowValues.Add(value.GetRawText());
                                break;
                        }
                    }
                    else
                    {
                        rowValues.Add(null);
                    }
                }

                cells.ImportObjectArray(rowValues.ToArray(), currentRow, 0, false);
                currentRow++;
            }
        }
        else
        {
            throw new NotSupportedException("Only .xml and .json files are supported.");
        }

        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        var transformer = new DataTransformer();

        string xmlInput = "data.xml";
        string jsonInput = "data.json";
        string output = "result.xlsx";

        if (File.Exists(xmlInput))
            transformer.Transform(xmlInput, output);
        else if (File.Exists(jsonInput))
            transformer.Transform(jsonInput, output);
        else
            Console.WriteLine("No input file found.");
    }
}