using System;
using Aspose.Cells;

public class GlobalizationResetDemo
{
    // Resets the workbook's globalization settings to the default English behavior.
    // Aspose.Cells applies globalization at the workbook level, so this affects all worksheets.
    public static void ResetGlobalizationSettings(Workbook workbook, params int[] worksheetIndices)
    {
        // Create a fresh instance of the default (English) globalization settings.
        GlobalizationSettings defaultSettings = new GlobalizationSettings();

        // Apply the default settings to the workbook.
        workbook.Settings.GlobalizationSettings = defaultSettings;
    }

    public static void Run()
    {
        // ---------- Create a workbook ----------
        Workbook wb = new Workbook();               // create new workbook
        wb.Worksheets.Add();                        // now we have two worksheets (indices 0 and 1)

        // ---------- Apply custom globalization to demonstrate the change ----------
        SettableGlobalizationSettings custom = new SettableGlobalizationSettings();
        custom.SetBooleanValueString(true, "TRUE_CUSTOM");
        custom.SetBooleanValueString(false, "FALSE_CUSTOM");
        wb.Settings.GlobalizationSettings = custom; // set custom settings

        // Populate some cells with boolean values.
        wb.Worksheets[0].Cells["A1"].PutValue(true);
        wb.Worksheets[0].Cells["A2"].PutValue(false);

        Console.WriteLine("Before reset:");
        Console.WriteLine($"Sheet0 A1: {wb.Worksheets[0].Cells["A1"].StringValue}");
        Console.WriteLine($"Sheet0 A2: {wb.Worksheets[0].Cells["A2"].StringValue}");

        // ---------- Reset globalization to default English for worksheet 0 ----------
        // (The reset is workbook‑wide; the worksheetIndices parameter is kept for API compatibility.)
        ResetGlobalizationSettings(wb, 0);

        Console.WriteLine("After reset:");
        Console.WriteLine($"Sheet0 A1: {wb.Worksheets[0].Cells["A1"].StringValue}");
        Console.WriteLine($"Sheet0 A2: {wb.Worksheets[0].Cells["A2"].StringValue}");

        // ---------- Save the workbook ----------
        wb.Save("GlobalizationResetDemo.xlsx");
    }
}

class Program
{
    static void Main()
    {
        GlobalizationResetDemo.Run();
    }
}