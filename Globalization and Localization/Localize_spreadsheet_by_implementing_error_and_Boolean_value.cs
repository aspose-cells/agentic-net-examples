using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings for Russian language
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values: TRUE -> ИСТИНА, FALSE -> ЛОЖЬ
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error values
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (replace with actual path)
            string inputPath = "input.xlsx";

            // Load the workbook (create/load lifecycle)
            Workbook workbook = new Workbook(inputPath);

            // Apply the custom Russian globalization settings
            workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

            // Example data: Boolean values and error strings
            Cells cells = workbook.Worksheets[0].Cells;
            cells[0, 0].PutValue(true);   // Boolean TRUE
            cells[0, 1].PutValue(false);  // Boolean FALSE

            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };

            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]); // Insert error strings
            }

            // Display localized values in console (optional verification)
            for (int i = 0; i < 9; i++)
            {
                Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
            }

            // Save the localized workbook (save lifecycle)
            string outputPath = "localized_output.xlsx";
            workbook.Save(outputPath);
        }
    }
}