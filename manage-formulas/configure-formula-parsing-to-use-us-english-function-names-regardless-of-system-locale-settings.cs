// Title: Force US‑English Function Names in Aspose.Cells Formula Parsing (C#/.NET)
// Description: Demonstrates how to assign a SettableGlobalizationSettings object with LocaleDependent set to false so that Aspose.Cells parses formulas using standard en‑US function names (e.g., SUM) regardless of the system locale, then calculates and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | formula parsing | US English functions | LocaleDependent false | SettableGlobalizationSettings | globalization settings | locale independent formulas | en-US function names
// Common Searches: Aspose.Cells use English function names on French Windows | disable locale dependent formulas Aspose.Cells .NET | set workbook globalization to en-US in Aspose.Cells | force SUM formula to parse in English regardless of OS language | how to ignore system locale for formulas in Aspose.Cells
// Developer Intent: Configure a workbook so that all formulas are interpreted with US‑English function names, independent of the machine’s regional settings.
// Use Cases: Run English‑written formulas on servers located in non‑English regions. | Create Excel files that behave consistently across different locale environments. | Share workbooks internationally without translating function names. | Automate calculations on cloud services where locale cannot be controlled.
// AI Prompts: Show C# code that sets Aspose.Cells GlobalizationSettings to enforce en‑US function names and then evaluates a formula. | Explain the effect of SettableGlobalizationSettings.LocaleDependent on formula parsing in Aspose.Cells. | Provide a step‑by‑step guide to disable locale‑dependent formula parsing and save the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaParsingDemo
{
    // Demonstrates how to assign a SettableGlobalizationSettings object with LocaleDependent set to false so that Aspose.Cells parses formulas using standard en‑US function names (e.g., SUM) regardless of the system locale, then calculates and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Create a SettableGlobalizationSettings instance.
            // This allows us to control how function names are interpreted.
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Ensure that formulas are parsed using the standard (en‑US) function names.
            // By default LocaleDependent is false, but we set it explicitly for clarity.
            // No mapping is required because we want the standard names.
            workbook.Settings.GlobalizationSettings = globalization;

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data.
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(15);
            sheet.Cells["A3"].PutValue(25);

            // Set a formula using the English function name "SUM".
            // Because LocaleDependent is false, the parser will treat "SUM" as the standard function
            // regardless of the machine's locale.
            sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

            // Calculate the workbook to evaluate the formula.
            workbook.CalculateFormula();

            // Output the result to verify that the formula was parsed correctly.
            Console.WriteLine("Result of =SUM(A1:A3): " + sheet.Cells["B1"].Value);

            // Save the workbook (uses the provided saving rule).
            workbook.Save("FormulaParsingUSEnglish.xlsx");
        }
    }
}
