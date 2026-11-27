// Title: How to replace the modify password of an existing Excel workbook with Aspose.Cells for .NET while keeping the data unchanged
// Common Searches: Aspose.Cells .NET change workbook modify password without affecting cell data | set new modify password on existing .xlsx using Aspose.Cells C# | update Excel file modify protection password programmatically with Aspose.Cells | replace modify password of loaded workbook Aspose.Cells reflection fallback | change Excel workbook modify password while preserving content in C#
// Tags: Aspose.Cells modify password API | C# set workbook modify password | Excel file protection password update | reflection fallback Aspose.Cells settings | preserve workbook content when changing password

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel workbook with Aspose.Cells, uses reflection to set a new modify password via the Settings.ModifyPassword property or the SetModifyPassword method (depending on the library version), and saves the file to a new location, leaving all worksheet data and formatting intact.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string newModifyPassword = "NewModifyPassword123";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Attempt to set the modify password using reflection to support multiple API versions
                try
                {
                    Type settingsType = workbook.Settings.GetType();

                    // Try property "ModifyPassword"
                    PropertyInfo prop = settingsType.GetProperty("ModifyPassword", BindingFlags.Public | BindingFlags.Instance);
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(workbook.Settings, newModifyPassword);
                        Console.WriteLine("Modify password set via property.");
                    }
                    else
                    {
                        // Fallback to method "SetModifyPassword"
                        MethodInfo method = settingsType.GetMethod("SetModifyPassword", BindingFlags.Public | BindingFlags.Instance);
                        if (method != null)
                        {
                            method.Invoke(workbook.Settings, new object[] { newModifyPassword });
                            Console.WriteLine("Modify password set via method.");
                        }
                        else
                        {
                            Console.WriteLine("Modify password feature is not available in this Aspose.Cells version.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set modify password: {ex.Message}");
                }

                // Save the workbook (overwrites the original file or saves to a new file)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
