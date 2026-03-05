using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to demonstrate localization of booleans, errors, and function names
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize boolean values (TRUE/FALSE)
        public override string GetBooleanValueString(bool value)
        {
            return value ? "VERDADERO" : "FALSO"; // Spanish representation
        }

        // Localize common Excel error strings
        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#DIV/0!" => "#DIV/0! (División por cero)",
                "#N/A"    => "#N/D (No disponible)",
                "#NAME?" => "#NOMBRE? (Nombre no reconocido)",
                "#REF!"  => "#REF! (Referencia inválida)",
                "#VALUE!"=> "#VALOR! (Valor incorrecto)",
                "#NUM!"  => "#NÚM! (Número no válido)",
                "#NULL!" => "#NULO! (Intersección nula)",
                _ => base.GetErrorValueString(err)
            };
        }

        // Localize the SUM function name
        public override string GetLocalFunctionName(string standardName)
        {
            return standardName.Equals("SUM", StringComparison.OrdinalIgnoreCase) ? "SUMA" : base.GetLocalFunctionName(standardName);
        }

        // Provide a custom default sheet name
        public override string GetDefaultSheetName()
        {
            return "HojaPersonalizada";
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with actual path if needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // Optionally set culture for loading (e.g., German uses comma as decimal separator)
            loadOptions.CultureInfo = new CultureInfo("de-DE");
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Demonstrate localized boolean and error values
            Cells cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue(true);   // Will display "VERDADERO"
            cells["A2"].PutValue(false);  // Will display "FALSO"
            cells["A3"].PutValue("#DIV/0!"); // Will display localized error string

            // Use the localized function name in a formula
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["B3"].Formula = "=SUMA(B1:B2)"; // "SUMA" is the localized name for SUM

            // Recalculate formulas to reflect the changes
            workbook.CalculateFormula();

            // Output some localized values to the console
            Console.WriteLine($"A1 (Boolean TRUE): {cells["A1"].StringValue}");
            Console.WriteLine($"A2 (Boolean FALSE): {cells["A2"].StringValue}");
            Console.WriteLine($"A3 (Error): {cells["A3"].StringValue}");
            Console.WriteLine($"B3 (Result of localized SUMA): {cells["B3"].StringValue}");

            // Save the localized workbook
            workbook.Save("localized_output.xlsx");
        }
    }
}