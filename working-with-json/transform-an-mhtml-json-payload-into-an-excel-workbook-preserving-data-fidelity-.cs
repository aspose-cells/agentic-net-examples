using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

public class MhtmlJsonToExcelConverter
{
    private class Payload
    {
        public string? MhtmlContent { get; set; }
    }

    public static void Convert(string jsonPayload, string excelOutputPath)
    {
        Payload? payload = JsonSerializer.Deserialize<Payload>(jsonPayload);
        if (payload == null || string.IsNullOrEmpty(payload.MhtmlContent))
            throw new ArgumentException("Invalid JSON payload or missing MHTML content.");

        string tempMhtPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".mht");
        File.WriteAllText(tempMhtPath, payload.MhtmlContent);

        try
        {
            Workbook workbook = new Workbook(tempMhtPath);
            workbook.Save(excelOutputPath, SaveFormat.Xlsx);
        }
        finally
        {
            if (File.Exists(tempMhtPath))
                File.Delete(tempMhtPath);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <program> <inputJsonPath> <outputExcelPath>");
            return;
        }

        string jsonPath = args[0];
        string excelPath = args[1];

        if (!File.Exists(jsonPath))
        {
            Console.WriteLine($"Input JSON file not found: {jsonPath}");
            return;
        }

        string json = File.ReadAllText(jsonPath);
        try
        {
            MhtmlJsonToExcelConverter.Convert(json, excelPath);
            Console.WriteLine($"Conversion successful. Excel saved to: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}