// Title: Copy all PageSetup settings from one worksheet to another using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy every writable PageSetup property from a source worksheet to a target worksheet. | Apply reflection to transfer page layout settings such as margins, orientation, and printable area between two sheets in an Excel workbook. | Programmatically duplicate the printing configuration of one worksheet onto another and save the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells copy page setup from one sheet to another C# | How to duplicate worksheet printing settings using reflection in .NET | Transfer margins and printable area between Excel worksheets with Aspose.Cells | Copy printable area orientation scaling from source sheet to destination sheet in C#
// Tags: copy worksheet page setup Aspose.Cells C# | transfer PageSetup properties via reflection | duplicate printable area margins Excel Aspose | clone worksheet printing configuration .NET | Aspose.Cells copy page layout settings

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel file, selects a source and a destination worksheet (with fallbacks), iterates over all readable and writable PageSetup properties using reflection, copies each value from the source sheet's PageSetup to the destination sheet's PageSetup, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the source worksheet; fallback to the first sheet if not found
            Worksheet sourceSheet = workbook.Worksheets["Source"] ?? workbook.Worksheets[0];
            if (sourceSheet == null)
                throw new InvalidOperationException("Source worksheet could not be determined.");

            // Retrieve the destination worksheet; fallback to the second sheet if not found
            Worksheet destinationSheet = workbook.Worksheets["Destination"] ?? 
                                         (workbook.Worksheets.Count > 1 ? workbook.Worksheets[1] : null);
            if (destinationSheet == null)
                throw new InvalidOperationException("Destination worksheet could not be determined.");

            // Get the PageSetup objects
            PageSetup srcSetup = sourceSheet.PageSetup;
            PageSetup destSetup = destinationSheet.PageSetup;

            // Copy writable PageSetup properties via reflection
            foreach (PropertyInfo prop in typeof(PageSetup).GetProperties())
            {
                if (prop.CanRead && prop.CanWrite && prop.GetIndexParameters().Length == 0)
                {
                    object value = prop.GetValue(srcSetup);
                    prop.SetValue(destSetup, value);
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
