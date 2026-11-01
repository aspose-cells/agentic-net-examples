// Title: How to unfreeze all panes in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove freeze panes from an Excel file with Aspose.Cells | Aspose.Cells unfreeze all worksheet panes programmatically | Call UnfreezePanes method in Aspose.Cells when it may be missing | Use reflection to disable FreezePanes property in older Aspose.Cells versions | Save an Excel workbook after clearing frozen panes with Aspose.Cells .NET
// Tags: Aspose.Cells UnfreezePanes API C# | Excel worksheet unfreeze panes .NET | fallback reflection FreezePanes property Aspose | programmatic removal of freeze panes Aspose.Cells | reset FreezePanes flag Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the first worksheet, and attempts to unfreeze all panes. It first invokes the UnfreezePanes method via reflection; if that method is unavailable, it falls back to setting the FreezePanes property to false using reflection, then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Attempt to unfreeze panes using the appropriate API for the current Aspose.Cells version
            try
            {
                // Prefer the UnfreezePanes method if it exists (newer versions)
                MethodInfo unfreezeMethod = sheet.GetType().GetMethod("UnfreezePanes", BindingFlags.Instance | BindingFlags.Public);
                if (unfreezeMethod != null)
                {
                    unfreezeMethod.Invoke(sheet, null);
                }
                else
                {
                    // Fallback: try to set a boolean FreezePanes property via reflection (older versions)
                    PropertyInfo freezeProp = sheet.GetType().GetProperty("FreezePanes", BindingFlags.Instance | BindingFlags.Public);
                    if (freezeProp != null && freezeProp.CanWrite)
                    {
                        freezeProp.SetValue(sheet, false);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log but continue – unfreeze failure should not stop saving
                Console.WriteLine($"Warning: Unable to unfreeze panes. {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
