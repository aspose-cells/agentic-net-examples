// Title: Set default worksheet font to Calibri 11 with Aspose.Cells for .NET
// Description: Creates a new Workbook, configures the default style to Calibri size 11, syncs Settings.DefaultStyleSettings and Workbook.DefaultStyle, optionally reapplies the style to existing cells, and saves the file as WorksheetDefaultFont.xlsx.
// Keywords: Aspose.Cells default font | set worksheet font .NET | Calibri 11 Aspose.Cells | default style settings | apply default style to cells | Aspose.Cells workbook font | C# Aspose.Cells example
// Common Searches: how to change default font in Aspose.Cells | Aspose.Cells set Calibri 11 for all worksheets | default style settings Aspose.Cells .NET example | apply default style to existing cells Aspose | C# code to set workbook default font Aspose.Cells
// Developer Intent: Configure a workbook so every new and existing cell uses Calibri 11 as the default font.
// Use Cases: Generate reports where the corporate font (Calibri 11) is applied automatically to every cell. | Standardize the appearance of legacy workbooks by syncing default style settings with existing cell formatting. | Create templates that enforce a consistent font without manually setting each cell style.
// AI Prompts: Write C# code using Aspose.Cells to set the workbook default font to Calibri 11 and propagate the change to all current cells. | Show how to update Settings.DefaultStyleSettings and Workbook.DefaultStyle for font name and size in Aspose.Cells. | Provide an example that re‑applies the default style to a worksheet after changing the default font in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new Workbook, configures the default style to Calibri size 11, syncs Settings.DefaultStyleSettings and Workbook.DefaultStyle, optionally reapplies the style to existing cells, and saves the file as WorksheetDefaultFont.xlsx.
public class SetWorksheetDefaultFont
{
    public static void Run()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the workbook's default style settings: Calibri, size 11
            // This influences all worksheets unless overridden
            workbook.Settings.DefaultStyleSettings.FontName = "Calibri";
            workbook.Settings.DefaultStyleSettings.FontSize = 11.0;

            // Also update the DefaultStyle object to keep it in sync
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // Apply the default style to the first worksheet (optional, ensures existing cells adopt it)
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells.ApplyStyle(workbook.CreateStyle(), new StyleFlag() { All = true });

            // Save the workbook (lifecycle: save)
            workbook.Save("WorksheetDefaultFont.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SetWorksheetDefaultFont.Run();
    }
}
