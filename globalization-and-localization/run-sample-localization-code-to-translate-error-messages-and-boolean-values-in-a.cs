using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings that translate boolean values and error messages to Russian
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Translate boolean values
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Translate error messages
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!":   return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A":    return "#Н/Д";
                case "#NUM!":   return "#ЧИСЛО!";
                case "#NULL!":  return "#ПУСТО!";
                default:        return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Prepare sample data: booleans and error strings
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            Cells cells = wb.Worksheets[0].Cells;

            // Put boolean values
            cells[0, 0].PutValue(true);
            cells[0, 1].PutValue(false);

            // Put error strings
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Apply custom globalization settings
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Display translated values in console
            for (int i = 0; i < 9; i++)
            {
                Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
            }

            // Save the workbook to a file
            wb.Save("LocalizedWorkbook.xlsx");
        }
    }
}