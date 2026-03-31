using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings for Russian language
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Override boolean display strings
        public override string GetBooleanValueString(bool value)
        {
            // TRUE -> ИСТИНА, FALSE -> ЛОЖЬ
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Override error value display strings
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
            // Create a new workbook
            Workbook wb = new Workbook();
            Cells cells = wb.Worksheets[0].Cells;

            // Populate cells with boolean values
            cells[0, 0].PutValue(true);
            cells[0, 1].PutValue(false);

            // Populate cells with standard error strings
            string[] errors = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Apply Russian globalization settings
            wb.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

            // Output localized values to console
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the workbook
            wb.Save("LocalizedWorkbook.xlsx");
        }
    }
}