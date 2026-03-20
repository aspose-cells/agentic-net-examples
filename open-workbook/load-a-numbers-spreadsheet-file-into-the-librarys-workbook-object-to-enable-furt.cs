using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Numbers;

class Program
{
    static void Main()
    {
        string filePath = "sample.numbers";

        Workbook workbook;

        if (File.Exists(filePath))
        {
            NumbersLoadOptions loadOptions = new NumbersLoadOptions
            {
                LoadTableType = LoadNumbersTableType.OneTablePerSheet
            };
            workbook = new Workbook(filePath, loadOptions);
        }
        else
        {
            workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Default");
        }

        Worksheet sheet = workbook.Worksheets[0];
        string cellValue = sheet.Cells["A1"].StringValue;
        Console.WriteLine($"Value of A1: {cellValue}");
    }
}