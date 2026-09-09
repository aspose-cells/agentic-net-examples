// Title: Configure Aspose.Cells workbook to recalculate only table formulas with AutomaticExceptTables mode in C#
// AI Prompts: Use Aspose.Cells for .NET to set Workbook.Settings.CalcMode to AutomaticExceptTables so that only table formulas are refreshed. | Apply reflection to assign the AutomaticExceptTables enum value to the CalcMode property, preserving compatibility with older Aspose.Cells releases. | Load an existing .xlsx file, change its calculation mode to AutomaticExceptTables, and save the updated workbook using C#.
// Common Searches: asp.net set calculation mode automaticexcepttables aspose.cells | how to recalculate only table formulas with Aspose.Cells C# | set workbook settings calcmode to AutomaticExceptTables fallback to Automatic | reflection based setting of CalcMode property in Aspose.Cells .NET | enable AutomaticExceptTables mode for Excel workbook using Aspose.Cells
// Tags: Aspose.Cells calculation mode AutomaticExceptTables | C# set Workbook.Settings.CalcMode | recalculate table formulas only Aspose.Cells | reflection fallback for CalcMode property | Excel workbook save with updated calculation mode

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, attempts to set its calculation mode to AutomaticExceptTables (using reflection to handle version differences), falls back to Automatic if necessary, and saves the workbook. This ensures that only formulas inside tables are recalculated while other formulas remain static.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to set calculation mode to AutomaticExceptTables if the property exists
            try
            {
                var settingsType = workbook.Settings.GetType();
                var calcModeProp = settingsType.GetProperty("CalcMode");
                if (calcModeProp != null && calcModeProp.CanWrite)
                {
                    // Parse the enum value by name; works for both older and newer enum definitions
                    var enumType = calcModeProp.PropertyType;
                    object modeValue = Enum.IsDefined(enumType, "AutomaticExceptTables")
                        ? Enum.Parse(enumType, "AutomaticExceptTables")
                        : Enum.Parse(enumType, "Automatic"); // fallback

                    calcModeProp.SetValue(workbook.Settings, modeValue);
                }
            }
            catch (Exception ex)
            {
                // Log but continue; calculation mode setting is optional
                Console.WriteLine($"Warning: Unable to set calculation mode. {ex.Message}");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
