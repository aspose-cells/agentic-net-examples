// Title: Enable the PivotTable ribbon UI programmatically in an Aspose.Cells workbook using C#
// Common Searches: C# code to show PivotTable ribbon after loading an Excel file with Aspose.Cells | How to turn on PivotTable UI elements in Aspose.Cells workbook settings | Using reflection to set EnablePivotTable property in Aspose.Cells .NET | Restore default PivotTable ribbon visibility in a saved workbook with Aspose.Cells | EnablePivotTable setting not available in older Aspose.Cells versions – workaround
// Tags: pivot table ribbon activation Aspose.Cells | workbook settings reflection .NET | default UI restoration Excel C# | Aspose.Cells UI customization | enable pivot UI via settings

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook (or creates a new one), uses reflection to set the Workbook.Settings.EnablePivotTable property when available, thereby programmatically enabling the PivotTable ribbon UI, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Attempt to enable PivotTable UI via reflection (property may not exist in older versions)
            try
            {
                var settingsType = workbook.Settings.GetType();
                var enablePivotProp = settingsType.GetProperty("EnablePivotTable");
                if (enablePivotProp != null && enablePivotProp.CanWrite)
                {
                    enablePivotProp.SetValue(workbook.Settings, true);
                }
            }
            catch (Exception ex)
            {
                // Log but continue if reflection fails
                Console.Error.WriteLine($"Warning: Unable to set EnablePivotTable – {ex.Message}");
            }

            // Save the workbook with the updated settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
