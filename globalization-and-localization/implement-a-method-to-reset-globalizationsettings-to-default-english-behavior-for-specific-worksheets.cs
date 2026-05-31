using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public static class GlobalizationResetHelper
    {
        /// <summary>
        /// Resets the workbook's globalization settings to the default English values.
        /// </summary>
        /// <param name="workbook">Workbook to modify.</param>
        /// <param name="worksheetIndices">Indices are ignored (kept for compatibility).</param>
        public static void ResetGlobalizationSettingsToDefaultEnglish(Workbook workbook, int[] worksheetIndices)
        {
            // Default English globalization settings
            GlobalizationSettings defaultSettings = new GlobalizationSettings();
            workbook.Settings.GlobalizationSettings = defaultSettings;
        }

        /// <summary>
        /// Demonstrates resetting globalization settings.
        /// </summary>
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Add sample data
                Worksheet ws1 = wb.Worksheets[0];
                ws1.Name = "Sheet1";
                ws1.Cells["A1"].PutValue(true);               // Boolean value
                ws1.Cells["A2"].PutValue("#DIV/0!");          // Error value

                // Apply custom Russian globalization
                SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
                customSettings.SetBooleanValueString(true, "ИСТИНА");
                customSettings.SetBooleanValueString(false, "ЛОЖЬ");
                wb.Settings.GlobalizationSettings = customSettings;

                // Show values before reset
                Console.WriteLine("Before reset:");
                Console.WriteLine($"A1 (bool) => {ws1.Cells["A1"].StringValue}");
                Console.WriteLine($"A2 (error) => {ws1.Cells["A2"].StringValue}");

                // Reset to default English (worksheetIndices not used)
                ResetGlobalizationSettingsToDefaultEnglish(wb, new int[] { 0 });

                // Show values after reset
                Console.WriteLine("\nAfter reset:");
                Console.WriteLine($"A1 (bool) => {ws1.Cells["A1"].StringValue}");
                Console.WriteLine($"A2 (error) => {ws1.Cells["A2"].StringValue}");

                // Save the workbook
                string outputPath = "ResetGlobalizationDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            GlobalizationResetHelper.Run();
        }
    }
}