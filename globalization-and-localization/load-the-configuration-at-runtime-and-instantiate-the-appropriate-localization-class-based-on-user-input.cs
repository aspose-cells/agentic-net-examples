using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Example locale – in real scenarios this could come from user input or a config file.
            string locale = "de";

            // Convert the locale string to a .NET CultureInfo instance.
            CultureInfo culture = GetCultureInfoFromLocale(locale);

            // Prepare load options with the desired culture.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = culture
            };

            Workbook workbook;

            // Ensure the input file exists; otherwise create a new workbook.
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                workbook = new Workbook(loadOptions);
            }

            // Apply globalization settings based on the locale.
            workbook.Settings.GlobalizationSettings = CreateLocalizationSettings(locale);

            // Example operation: write the locale display name into a cell.
            workbook.Worksheets[0].Cells["A1"].PutValue($"Locale: {culture.DisplayName}");

            // Save the workbook.
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Maps simple locale codes to full .NET culture identifiers.
    static CultureInfo GetCultureInfoFromLocale(string locale)
    {
        return locale switch
        {
            "ar" => new CultureInfo("ar-SA"),
            "bg" => new CultureInfo("bg-BG"),
            "ca" => new CultureInfo("ca-ES"),
            "cs" => new CultureInfo("cs-CZ"),
            "da" => new CultureInfo("da-DK"),
            "de" => new CultureInfo("de-DE"),
            "el" => new CultureInfo("el-GR"),
            "en" => new CultureInfo("en-US"),
            "es" => new CultureInfo("es-ES"),
            "fa" => new CultureInfo("fa-IR"),
            "fr" => new CultureInfo("fr-FR"),
            "he" => new CultureInfo("he-IL"),
            "hi" => new CultureInfo("hi-IN"),
            "hr" => new CultureInfo("hr-HR"),
            "hu" => new CultureInfo("hu-HU"),
            "id" => new CultureInfo("id-ID"),
            "it" => new CultureInfo("it-IT"),
            "ja" => new CultureInfo("ja-JP"),
            "ka" => new CultureInfo("ka-GE"),
            "ko" => new CultureInfo("ko-KR"),
            "lt" => new CultureInfo("lt-LT"),
            "ms" => new CultureInfo("ms-MY"),
            "nl" => new CultureInfo("nl-NL"),
            "pl" => new CultureInfo("pl-PL"),
            "pt" => new CultureInfo("pt-PT"),
            "ro" => new CultureInfo("ro-RO"),
            "ru" => new CultureInfo("ru-RU"),
            "sk" => new CultureInfo("sk-SK"),
            "th" => new CultureInfo("th-TH"),
            "tr" => new CultureInfo("tr-TR"),
            "uk" => new CultureInfo("uk-UA"),
            "vi" => new CultureInfo("vi-VN"),
            "zh" => new CultureInfo("zh-CN"),
            "zh-hant" => new CultureInfo("zh-TW"),
            _ => CultureInfo.InvariantCulture,
        };
    }

    // Returns a GlobalizationSettings instance. Adjust as needed for specific locales.
    static GlobalizationSettings CreateLocalizationSettings(string locale)
    {
        // For demonstration, we return a default GlobalizationSettings instance.
        // Customization per locale can be added here if required.
        return new GlobalizationSettings();
    }
}