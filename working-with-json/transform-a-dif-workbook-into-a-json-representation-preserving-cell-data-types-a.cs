using System;
using Aspose.Cells;

namespace DifToJsonExample
{
    class Program
    {
        static void Main()
        {
            string difFilePath = "source.dif";

            Workbook workbook = new Workbook(difFilePath);
            Worksheet worksheet = workbook.Worksheets[0];

            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            string jsonResult;

            if (maxRow >= 0 && maxCol >= 0)
            {
                Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    ToExcelStruct = true,
                    ExportEmptyCells = true,
                    Indent = "  "
                };
                jsonResult = usedRange.ToJson(jsonOptions);
            }
            else
            {
                jsonResult = "[]";
            }

            Console.WriteLine(jsonResult);
        }
    }
}