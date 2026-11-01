// Title: How to disable all error‑checking options in an Aspose.Cells workbook with C# (using reflection)
// Common Searches: C# Aspose.Cells disable all error checking in a workbook | using reflection to turn off Excel validation errors with Aspose.Cells .NET | set EnableErrorChecking false in Aspose.Cells workbook programmatically | disable formula and data validation error checks in Aspose.Cells C# | remove Excel error indicators from generated workbook using Aspose.Cells | Aspose.Cells workbook Settings ErrorCheckingOptions disable via code
// Tags: Aspose.Cells disable error checking via reflection | C# workbook Settings.EnableErrorChecking false | Aspose.Cells ErrorCheckingOptions property manipulation | turn off Excel validation errors in .NET workbook | disable formula error indicators Aspose.Cells | programmatic suppression of Excel error checks C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example defines a WorksheetHelper.DisableAllErrorChecks method that accepts a Workbook, uses reflection to set Settings.EnableErrorChecking to false and iterates through the Settings.ErrorCheckingOptions properties, disabling each specific check such as EnableNumberStoredAsText, EnableFormula, and EnableDataValidation. It gracefully handles missing properties and throws a clear exception on failure. The Program class shows loading or creating a workbook, invoking the helper to turn off all error checks, and saving the modified file.
public static class WorksheetHelper
{
    /// <param name="workbook">The workbook to modify.</param>
    public static void DisableAllErrorChecks(Workbook workbook)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));

        try
        {
            // Attempt to disable global error checking if the property exists.
            PropertyInfo enableErrorProp = workbook.Settings.GetType().GetProperty("EnableErrorChecking");
            if (enableErrorProp != null && enableErrorProp.CanWrite)
            {
                enableErrorProp.SetValue(workbook.Settings, false);
            }

            // Retrieve ErrorCheckingOptions via reflection (API may vary between versions).
            PropertyInfo optionsProp = workbook.Settings.GetType().GetProperty("ErrorCheckingOptions");
            if (optionsProp != null)
            {
                object options = optionsProp.GetValue(workbook.Settings);
                if (options != null)
                {
                    // Helper to set boolean properties safely.
                    void SetOption(string name, bool val)
                    {
                        PropertyInfo pi = options.GetType().GetProperty(name);
                        if (pi != null && pi.CanWrite)
                        {
                            pi.SetValue(options, val);
                        }
                    }

                    SetOption("EnableNumberStoredAsText", false);
                    SetOption("EnableFormula", false);
                    SetOption("EnableInconsistentFormula", false);
                    SetOption("EnableEmptyCellReference", false);
                    SetOption("EnableListDataValidation", false);
                    SetOption("EnableOmittedCells", false);
                    SetOption("EnableUnlockedFormula", false);
                    SetOption("EnableDataValidation", false);
                    SetOption("EnableFormulaRange", false);
                    SetOption("EnableFormulaError", false);
                    SetOption("EnableTextDate", false);
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to disable error checks on the workbook.", ex);
        }
    }
}

public class Program
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Disable all error checks.
            WorksheetHelper.DisableAllErrorChecks(workbook);

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
