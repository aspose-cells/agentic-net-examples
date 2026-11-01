// Title: Hide zero values on worksheets whose names start with "Q" using Aspose.Cells WorksheetCollection in C#
// Common Searches: C# Aspose.Cells hide zero values on worksheets with name prefix Q | How to set IsZeroValuesHidden for specific sheets using WorksheetCollection in Aspose.Cells | Iterate worksheets and hide zero values based on sheet name in Aspose.Cells .NET | Apply reflection to set IsZeroValuesHidden property for Excel sheets starting with Q in C# | Filter worksheets by name and hide zero values with Aspose.Cells API
// Tags: WorksheetCollection filter by name prefix | set IsZeroValuesHidden property Aspose.Cells | hide zero values Excel worksheet C# | process worksheets starting with Q Aspose.Cells | reflection based property setting Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through all worksheets, and for each sheet whose name begins with "Q" it uses reflection to set the IsZeroValuesHidden property to true via the WorksheetCollection API, effectively hiding zero values before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process only worksheets whose name starts with "Q"
                    if (sheet.Name.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
                    {
                        // Hide zero values on this worksheet (using reflection for compatibility)
                        PropertyInfo prop = sheet.GetType().GetProperty("IsZeroValuesHidden", BindingFlags.Public | BindingFlags.Instance);
                        if (prop != null && prop.CanWrite)
                        {
                            prop.SetValue(sheet, true);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
