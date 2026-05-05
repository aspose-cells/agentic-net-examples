using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace JsonToExcelConverter
{
    public class WorkbookJson
    {
        public List<WorksheetJson> Worksheets { get; set; } = new();
    }

    public class WorksheetJson
    {
        public string Name { get; set; } = "";
        public List<RowJson> Rows { get; set; } = new();
    }

    public class RowJson
    {
        public List<object> Cells { get; set; } = new();
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: JsonToExcelConverter <input.json> <output.xlsx>");
                return;
            }

            string jsonPath = args[0];
            string excelPath = args[1];

            string json = File.ReadAllText(jsonPath);
            var workbookData = JsonSerializer.Deserialize<WorkbookJson>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            var workbook = new Workbook();
            workbook.Worksheets.Clear();

            foreach (var wsJson in workbookData.Worksheets)
            {
                var sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet.Name = wsJson.Name;

                for (int i = 0; i < wsJson.Rows.Count; i++)
                {
                    var row = wsJson.Rows[i];
                    for (int j = 0; j < row.Cells.Count; j++)
                    {
                        var cell = sheet.Cells[i, j];
                        var value = row.Cells[j];

                        if (value is JsonElement je)
                        {
                            switch (je.ValueKind)
                            {
                                case JsonValueKind.Number:
                                    if (je.TryGetInt64(out long l))
                                        cell.PutValue(l);
                                    else if (je.TryGetDouble(out double d))
                                        cell.PutValue(d);
                                    break;
                                case JsonValueKind.String:
                                    cell.PutValue(je.GetString());
                                    break;
                                case JsonValueKind.True:
                                case JsonValueKind.False:
                                    cell.PutValue(je.GetBoolean());
                                    break;
                                default:
                                    cell.PutValue(je.ToString());
                                    break;
                            }
                        }
                        else
                        {
                            cell.PutValue(value);
                        }
                    }
                }
            }

            workbook.Save(excelPath);
            Console.WriteLine($"Excel file saved to {excelPath}");
        }
    }
}