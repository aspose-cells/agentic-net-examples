using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook wb = new Workbook(inputPath);

        // Prepare sample data: boolean values and error strings
        string[] errors = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
        Cells cells = wb.Worksheets[0].Cells;

        cells[0, 0].PutValue(true);
        cells[0, 1].PutValue(false);
        for (int i = 0; i < errors.Length; i++)
        {
            cells[0, i + 2].PutValue(errors[i]);
        }

        // Apply custom globalization settings for Boolean and error localization
        wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Display localized string values of the cells
        for (int i = 0; i < 9; i++)
        {
            Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
        }

        // Save the modified workbook
        wb.Save("output.xlsx");
    }

    // Custom globalization settings overriding default Boolean and error strings
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?": return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!": return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A": return "#Н/Д";
                case "#NUM!": return "#ЧИСЛО!";
                case "#NULL!": return "#ПУСТО!";
                default: return base.GetErrorValueString(err);
            }
        }
    }
}