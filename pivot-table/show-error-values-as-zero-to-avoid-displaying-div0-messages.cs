// Title: Replace all Excel error values (e.g., #DIV/0!) with zero using Aspose.Cells GlobalizationSettings in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, inserts a formula causing a division‑by‑zero error, and applies a custom GlobalizationSettings subclass so that any error is displayed as "0" before saving the file. | Write a C# class inheriting from Aspose.Cells.GlobalizationSettings that overrides GetErrorValueString to return "0", and demonstrate how to assign this class to Workbook.Settings.GlobalizationSettings. | Show how to calculate formulas in an Aspose.Cells workbook and save it as an .xlsx file where cells containing errors are rendered as zero.
// Common Searches: how to show zero instead of #DIV/0! in Excel files generated with Aspose.Cells C# | Aspose.Cells custom GlobalizationSettings to replace all error values with 0 | C# Aspose.Cells example for overriding GetErrorValueString for error handling | save workbook with error cells displayed as 0 using Aspose.Cells | replace Excel error strings with numeric zero using Aspose.Cells API
// Tags: Aspose.Cells error value globalization | C# custom GetErrorValueString implementation | replace Excel error strings with zero | save workbook with zeroed error cells | globalization settings for error handling

using System;
using Aspose.Cells;

// The example creates an Aspose.Cells workbook, adds a formula that generates a division‑by‑zero error, calculates the formulas, applies a custom GlobalizationSettings class that returns "0" for any error via GetErrorValueString, and saves the workbook so error cells display zero.
public class Program
{
    public static void Main()
    {
        try
        {
            ShowErrorAsZero.Run();
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class ShowErrorAsZero
{
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Cell with a formula that generates a division by zero error
        Cell errorCell = worksheet.Cells["A1"];
        errorCell.Formula = "=1/0";

        // Calculate formulas so the error value is produced
        workbook.CalculateFormula();

        // Apply custom globalization settings to replace any error string with "0"
        workbook.Settings.GlobalizationSettings = new ZeroErrorGlobalizationSettings();

        // Save the workbook; cells with errors will now display 0
        string outputPath = "ErrorAsZero.xlsx";
        workbook.Save(outputPath);
    }
}

// Custom globalization settings that maps all error values to zero
public class ZeroErrorGlobalizationSettings : GlobalizationSettings
{
    public override string GetErrorValueString(string err)
    {
        // Return "0" for any error (e.g., #DIV/0!, #VALUE!, etc.)
        return "0";
    }
}
