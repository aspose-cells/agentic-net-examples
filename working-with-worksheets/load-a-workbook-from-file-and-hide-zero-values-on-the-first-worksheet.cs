// Title: Load an Excel workbook from a file and hide zero values on the first worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load existing workbook and hide zero values on first sheet | how to set Worksheet.IsZeroValuesHidden property in Aspose.Cells .NET | C# hide zero values in Excel worksheet using Aspose.Cells reflection | load workbook from file and hide zero values with Aspose.Cells API | Aspose.Cells hide zero values when property is unavailable in older versions
// Tags: load workbook from file Aspose.Cells | set Worksheet.IsZeroValuesHidden via reflection | first worksheet zero-value hiding Aspose.Cells | save modified workbook Aspose.Cells | handle missing input.xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads "input.xlsx" (or creates a new workbook if the file is absent), accesses the first worksheet, uses reflection to enable the Worksheet.IsZeroValuesHidden property when available, and saves the result as "output.xlsx" while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide zero values if the property is available in the current Aspose.Cells version
            var prop = typeof(Worksheet).GetProperty("IsZeroValuesHidden");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(worksheet, true);
            }

            // Save the workbook with the changes
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
