// Title: Hide PivotTable field list and toolbar in an Excel workbook at runtime using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells hide PivotTable field list at runtime | disable PivotTable toolbar in Excel file with Aspose.Cells .NET | remove PivotTable UI elements from workbook using Aspose.Cells | how to suppress PivotTable ribbon controls programmatically in C# | Aspose.Cells set ShowPivotTableToolbar false via reflection
// Tags: pivot table field list visibility Aspose.Cells | pivot table toolbar visibility Aspose.Cells | runtime workbook UI customization C# | Aspose.Cells reflection for settings properties | Excel workbook hide pivot UI .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, uses reflection to set the Settings properties ShowPivotTableFieldList and ShowPivotTableToolbar to false when they exist, and saves the file, effectively hiding the PivotTable field list and toolbar during runtime.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Attempt to hide the PivotTable field list if the property exists in the current Aspose.Cells version
            var settings = workbook.Settings;
            var fieldListProp = settings.GetType().GetProperty("ShowPivotTableFieldList");
            if (fieldListProp != null && fieldListProp.CanWrite)
            {
                fieldListProp.SetValue(settings, false);
            }

            // Attempt to hide the PivotTable toolbar if the property exists
            var toolbarProp = settings.GetType().GetProperty("ShowPivotTableToolbar");
            if (toolbarProp != null && toolbarProp.CanWrite)
            {
                toolbarProp.SetValue(settings, false);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
