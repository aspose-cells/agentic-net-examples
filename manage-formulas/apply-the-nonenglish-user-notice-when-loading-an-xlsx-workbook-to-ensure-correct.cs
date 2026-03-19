using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            LoadWithNonEnglishUserNotice.Run();
        }
    }

    public class LoadWithNonEnglishUserNotice
    {
        public static void Run()
        {
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguageCode = CountryCode.Germany;

            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Loaded with German language code");

            workbook.Save("output.xlsx");
        }
    }
}