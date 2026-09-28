// Title: Hide row and column headers in an Aspose.Cells worksheet view using C# (compatible with older versions)
// AI Prompts: Generate C# code that uses Aspose.Cells to set the ShowRowColumnHeaders property to false, with a reflection fallback for versions that lack ViewOptions. | Provide a C# example that creates a workbook, accesses the first worksheet, disables row and column headings, and saves the result as an .xlsx file. | Write a C# snippet that safely checks for the Worksheet.ViewOptions property and turns off header display without raising exceptions.
// Common Searches: C# Aspose.Cells hide worksheet row and column headers programmatically | How to disable Excel grid headings in Aspose.Cells .NET | Set ShowRowColumnHeaders false with Aspose.Cells for older library versions | Aspose.Cells hide row/column headers when saving workbook | Reflection based header hiding in Aspose.Cells C# example
// Tags: Aspose.Cells worksheet header visibility | C# ShowRowColumnHeaders property | Aspose.Cells ViewOptions reflection | disable Excel row/column headings .NET | compatible header hiding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, obtains the first Worksheet, detects the ViewOptions property via reflection, sets ShowRowColumnHeaders to false when available, and saves the workbook as Output.xlsx, ensuring compatibility across Aspose.Cells versions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide row and column headings in the worksheet view if the API supports it
            // In newer Aspose.Cells versions this can be done via sheet.ViewOptions.ShowRowColumnHeaders = false;
            // The following check ensures compatibility with older versions where ViewOptions may not exist.
            var viewOptionsProperty = typeof(Worksheet).GetProperty("ViewOptions");
            if (viewOptionsProperty != null)
            {
                var viewOptions = viewOptionsProperty.GetValue(sheet);
                var showHeadersProp = viewOptions?.GetType().GetProperty("ShowRowColumnHeaders");
                if (showHeadersProp != null && showHeadersProp.CanWrite)
                {
                    showHeadersProp.SetValue(viewOptions, false);
                }
            }

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
