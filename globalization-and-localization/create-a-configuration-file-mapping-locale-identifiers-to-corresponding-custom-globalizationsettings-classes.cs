using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings for Russian locale
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    // Custom globalization settings for French locale
    public class FrenchGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "VRAI" : "FAUX";
        }

        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#NAME?" => "#NOM?",
                "#DIV/0!" => "#DIV/0!",
                "#REF!" => "#REF!",
                "#VALUE!" => "#VALEUR!",
                "#N/A" => "#N/D",
                "#NUM!" => "#NUM!",
                "#NULL!" => "#VIDE!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    // Factory that returns the appropriate GlobalizationSettings based on LCID
    public static class GlobalizationSettingsFactory
    {
        // Mapping of locale identifiers (LCID) to custom settings instances
        private static readonly Dictionary<int, GlobalizationSettings> _localeMap = new()
        {
            { 1049, new RussianGlobalizationSettings() }, // Russian - Russia
            { 1036, new FrenchGlobalizationSettings() }   // French - France
            // Add more mappings as needed
        };

        public static GlobalizationSettings GetSettings(int lcid)
        {
            // Return the custom settings if found; otherwise return default settings
            return _localeMap.TryGetValue(lcid, out var settings) ? settings : new GlobalizationSettings();
        }
    }

    class Program
    {
        static void Main()
        {
            // Example LCID to use; change this value to test different locales
            int localeId = 1049; // Russian

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data with boolean and error values
            cells[0, 0].PutValue(true);
            cells[0, 1].PutValue(false);
            string[] errors = { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Apply the custom globalization settings based on the locale identifier
            workbook.Settings.GlobalizationSettings = GlobalizationSettingsFactory.GetSettings(localeId);

            // Demonstrate that the settings affect cell string values
            for (int i = 0; i < 9; i++)
            {
                Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
            }

            // Save the workbook
            workbook.Save("LocalizedWorkbook.xlsx");
        }
    }
}