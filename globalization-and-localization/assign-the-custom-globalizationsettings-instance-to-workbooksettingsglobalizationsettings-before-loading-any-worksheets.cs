using System;
using Aspose.Cells;

// Custom globalization settings derived from GlobalizationSettings
class CustomGlobalizationSettings : GlobalizationSettings
{
    // Override boolean display strings
    public override string GetBooleanValueString(bool value)
    {
        return value ? "ИСТИНА" : "ЛОЖЬ";
    }

    // Override error value strings
    public override string GetErrorValueString(string error)
    {
        return error switch
        {
            "#NAME?"   => "#ИМЯ?",
            "#DIV/0!" => "#ДЕЛ/0!",
            "#REF!"   => "#ССЫЛКА!",
            "#VALUE!" => "#ЗНАЧ!",
            "#N/A"    => "#Н/Д",
            "#NUM!"   => "#ЧИСЛО!",
            "#NULL!"  => "#ПУСТО!",
            _         => base.GetErrorValueString(error)
        };
    }
}

class Program
{
    static void Main()
    {
        // Create LoadOptions (no special options needed for this example)
        LoadOptions loadOptions = new LoadOptions();

        // Load the workbook; at this point no worksheet has been accessed yet
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Assign the custom globalization settings BEFORE any worksheet is accessed
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Now worksheets can be used safely with the custom settings applied
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        Console.WriteLine($"Cell[0,0] value: {cell.StringValue}");

        // Save the workbook if needed
        workbook.Save("output.xlsx");
    }
}