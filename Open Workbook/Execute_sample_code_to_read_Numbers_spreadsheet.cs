using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Numbers;

class ReadNumbersExample
{
    static void Main()
    {
        string numbersPath = "sample.numbers";

        NumbersLoadOptions loadOptions = new NumbersLoadOptions
        {
            LoadTableType = LoadNumbersTableType.OneTablePerSheet,
            CultureInfo = new CultureInfo("en-US")
        };

        Workbook workbook;

        if (File.Exists(numbersPath))
        {
            workbook = new Workbook(numbersPath, loadOptions);
        }
        else
        {
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B2"].PutValue(12345);
        }

        Worksheet firstSheet = workbook.Worksheets[0];
        Console.WriteLine("First worksheet name: " + firstSheet.Name);
        Console.WriteLine("Cell A1 value: " + firstSheet.Cells["A1"].StringValue);

        workbook.Save("output.xlsx");
    }
}