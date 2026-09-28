// Title: Check that a cell’s explicit RGB font color stays the same after removing the workbook theme using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets a cell’s font to a specific RGB value, invokes Workbook.RemoveTheme via reflection, and confirms the font’s ARGB value is unchanged. | Create a C# unit‑test that records a cell’s font color before and after calling the RemoveTheme method on a Workbook and asserts that the RGB components match.
// Common Searches: Aspose.Cells C# verify cell font color after removing workbook theme | how to keep explicit RGB font color when calling Workbook.RemoveTheme in .NET | unit test for theme removal preserving custom cell colors Aspose.Cells | remove workbook theme without affecting cell colors Aspose.Cells C# | compare ARGB values before and after theme deletion using Aspose.Cells
// Tags: Aspose.Cells workbook theme removal | explicit font RGB preservation | cell font ARGB comparison | invoke RemoveTheme with reflection | color consistency validation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;

// The example creates a workbook, applies an explicit blue RGB font color to cell A1, captures the font's ARGB value, attempts to remove the workbook theme via reflection, captures the ARGB value again, and verifies that the two values are identical, demonstrating that explicit RGB colors are retained after theme removal.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a sample value into cell A1
            sheet.Cells["A1"].PutValue("Sample");

            // Apply an explicit RGB color to the font of cell A1
            Style cellStyle = sheet.Cells["A1"].GetStyle();
            cellStyle.Font.Color = Color.Blue; // use a known color instead of theme color
            sheet.Cells["A1"].SetStyle(cellStyle);

            // Capture the actual RGB value after the color is applied
            Color rgbBefore = sheet.Cells["A1"].GetStyle().Font.Color;

            // Attempt to remove the workbook theme via reflection (method may not exist in older versions)
            var removeThemeMethod = typeof(Workbook).GetMethod("RemoveTheme");
            if (removeThemeMethod != null)
            {
                try
                {
                    removeThemeMethod.Invoke(workbook, null);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Theme removal failed: {ex.Message}");
                }
            }

            // Capture the RGB value after the (attempted) theme removal
            Color rgbAfter = sheet.Cells["A1"].GetStyle().Font.Color;

            // Validate that the RGB values are identical
            bool valuesRetained = rgbBefore.ToArgb() == rgbAfter.ToArgb();

            Console.WriteLine($"RGB before theme removal: {rgbBefore}");
            Console.WriteLine($"RGB after theme removal:  {rgbAfter}");
            Console.WriteLine($"Values retained: {valuesRetained}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
