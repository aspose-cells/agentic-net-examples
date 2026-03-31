using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook (required lifecycle step)
        Workbook workbook = new Workbook();

        // Map each ValidationType to its description
        var descriptions = new Dictionary<ValidationType, string>
        {
            { ValidationType.AnyValue,   "Any value validation type." },
            { ValidationType.WholeNumber, "Whole number validation type." },
            { ValidationType.Decimal,    "Decimal validation type." },
            { ValidationType.List,       "List validation type." },
            { ValidationType.Date,       "Date validation type." },
            { ValidationType.Time,       "Time validation type." },
            { ValidationType.TextLength, "Text length validation type." },
            { ValidationType.Custom,     "Custom validation type." }
        };

        // Enumerate and display all supported data validation types
        Console.WriteLine("Supported Data Validation Types for XLSX worksheets:");
        foreach (ValidationType vt in Enum.GetValues(typeof(ValidationType)))
        {
            string description = descriptions.ContainsKey(vt) ? descriptions[vt] : "No description available.";
            Console.WriteLine($"{(int)vt}. {vt} - {description}");
        }

        // Save the workbook (optional, demonstrates save lifecycle)
        workbook.Save("DataValidationTypes.xlsx");
    }
}