// Title: Disable Automatic Formula Recalculation in Aspose.Cells for .NET (C#)
// Description: Learn how to set Aspose.Cells workbook settings to Manual calculation mode and turn off CalculateOnSave, so formulas stay unevaluated until you call CalculateFormula explicitly.
// Keywords: Aspose.Cells manual calculation mode | disable auto formula recalculation C# | CalcModeType.Manual Aspose.Cells | prevent formula evaluation on save | Aspose.Cells workbook settings | control formula calculation .NET | performance optimization large spreadsheets | Aspose.Cells CalculateOnSave false
// Common Searches: Aspose.Cells turn off automatic calculation C# | set workbook to manual calculation mode Aspose | disable formula recalculation on save Aspose.Cells | how to prevent formulas from recalculating in .NET | Aspose.Cells performance manual calc mode
// Developer Intent: Stop automatic formula evaluation so calculations run only when explicitly triggered.
// Use Cases: Generate a spreadsheet with thousands of formulas, defer calculation to improve data‑entry speed. | Create a template where formulas must remain untouched until a later processing step. | Export data from an application, save the workbook, and let the end user decide when to recalculate.
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, sets CalculationMode to Manual, disables CalculateOnSave, adds sample data with a SUM formula, and saves the file. | Show how to switch a workbook back to automatic calculation after performing custom updates with Aspose.Cells. | Provide an example that reads a workbook saved in manual mode, invokes CalculateFormula, and then saves the evaluated result.

using System;
using System.IO;
using Aspose.Cells;

// Learn how to set Aspose.Cells workbook settings to Manual calculation mode and turn off CalculateOnSave, so formulas stay unevaluated until you call CalculateFormula explicitly.
public class DisableAutoRecalcDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set calculation mode to Manual to prevent automatic recalculation
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Disable recalculation when the workbook is saved
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Add sample data and a formula (the formula will not be evaluated automatically)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Determine output path and ensure directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ManualCalc.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook; the formula remains unevaluated until CalculateFormula is called explicitly
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook processing: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            DisableAutoRecalcDemo.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
