using System;
using Aspose.Cells;

namespace AsposeCellsErrorAndBooleanLocalization
{
    // Custom globalization settings to localize Boolean and error strings
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values (e.g., Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error strings (example translations)
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

    class Program
    {
        static void Main()
        {
            // Configure load options: skip data validation checks and formula parsing on open
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CheckDataValid = false;
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook from an existing XLSX file
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Apply custom globalization settings for localization
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check for error values and output localized error string
                    if (cell.IsErrorValue)
                    {
                        Console.WriteLine($"Error at {cell.Name}: {cell.StringValue}");
                    }
                    // Check for Boolean values and output localized Boolean string
                    else if (cell.Value is bool boolVal)
                    {
                        string localizedBool = workbook.Settings.GlobalizationSettings.GetBooleanValueString(boolVal);
                        Console.WriteLine($"Boolean at {cell.Name}: {localizedBool}");
                    }
                }
            }

            // Save the workbook (the localization affects only string representations, not the stored values)
            workbook.Save("output.xlsx");
        }
    }
}