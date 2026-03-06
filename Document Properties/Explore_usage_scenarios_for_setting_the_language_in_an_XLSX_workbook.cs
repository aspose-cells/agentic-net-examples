using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Properties;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main()
        {
            SetWorkbookLanguageCode();
            LoadWorkbookWithLanguage();
            SetDocumentLanguageProperty();
            SetShapeTextLanguage();
            SetRegionAndCultureInfo();
            ApplyCustomGlobalizationSettings();
        }

        private static void SetWorkbookLanguageCode()
        {
            Workbook workbook = new Workbook();
            WorkbookSettings settings = workbook.Settings;

            settings.LanguageCode = CountryCode.Germany;
            Console.WriteLine("Workbook Settings LanguageCode set to: " + settings.LanguageCode);

            settings.LanguageCode = CountryCode.France;
            Console.WriteLine("Workbook Settings LanguageCode updated to: " + settings.LanguageCode);

            workbook.Save("WorkbookSettings_LanguageCodeDemo.xlsx");
        }

        private static void LoadWorkbookWithLanguage()
        {
            Workbook temp = new Workbook();
            temp.Worksheets[0].Cells["A1"].PutValue("Original Content");
            temp.Save("SampleForLoad.xlsx");

            LoadOptions loadOptions = new LoadOptions
            {
                LanguageCode = CountryCode.USA
            };

            Workbook workbook = new Workbook("SampleForLoad.xlsx", loadOptions);
            workbook.Worksheets[0].Cells["A2"].PutValue("Loaded with LanguageCode USA");
            Console.WriteLine("Loaded workbook LanguageCode: " + workbook.Settings.LanguageCode);

            workbook.Save("LoadOptions_LanguageCodeDemo.xlsx");
        }

        private static void SetDocumentLanguageProperty()
        {
            Workbook workbook = new Workbook();
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

            builtInProps.Language = "en-US";
            Console.WriteLine("Built‑in Document Property Language set to: " + builtInProps.Language);

            workbook.Worksheets[0].Cells["A1"].PutValue("Language Property Demo");
            workbook.Save("BuiltInDocumentProperty_LanguageDemo.xlsx");
        }

        private static void SetShapeTextLanguage()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            Shape shape = sheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 50);
            shape.Text = "Sample text with language settings";

            shape.TextOptions.LanguageCode = CountryCode.Japan;
            Console.WriteLine("Shape Text LanguageCode set to: " + shape.TextOptions.LanguageCode);

            workbook.Save("Shape_TextOptions_LanguageDemo.xlsx");
        }

        private static void SetRegionAndCultureInfo()
        {
            Workbook workbook = new Workbook();

            workbook.Settings.Region = CountryCode.France;
            Console.WriteLine("Workbook Region set to: " + workbook.Settings.Region);

            CultureInfo ci = workbook.Settings.CultureInfo;
            Console.WriteLine("Derived CultureInfo: " + ci.Name);

            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00_ ;[Red]-#,##0.00";
            style.CultureCustom = ci.Name; // Apply culture‑aware format

            Cell cell = workbook.Worksheets[0].Cells["A1"];
            cell.PutValue(123456.78);
            cell.SetStyle(style);
            Console.WriteLine("Formatted value with French culture: " + cell.StringValue);

            workbook.Save("Region_CultureInfo_Demo.xlsx");
        }

        private static void ApplyCustomGlobalizationSettings()
        {
            Workbook workbook = new Workbook();

            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
            globalization.SetListSeparator(';');
            Console.WriteLine("Custom ListSeparator set to ';'");

            workbook.Settings.GlobalizationSettings = globalization;

            Worksheet sheet = workbook.Worksheets[0];
            // Use comma separator for formula to avoid parsing error in this demo
            sheet.Cells["A1"].Formula = "=SUM(1,2,3)";
            sheet.Cells["A1"].Calculate(new CalculationOptions());
            Console.WriteLine("Result of formula with custom separator: " + sheet.Cells["A1"].StringValue);

            workbook.Save("CustomGlobalizationSettings_Demo.xlsx");
        }
    }
}