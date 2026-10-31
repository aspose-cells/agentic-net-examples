// Title: Copy worksheet PageSetup configuration from a source sheet to a destination sheet using Aspose.Cells for .NET (C#)
// Common Searches: aspnet copy page setup from one worksheet to another using Aspose.Cells | c# transfer printable page settings between Excel sheets with Aspose.Cells | duplicate worksheet printing configuration programmatically in .NET | reflectively copy all writable PageSetup properties Aspose.Cells C# | how to copy page margins and orientation between worksheets using Aspose.Cells
// Tags: Aspose.Cells copy worksheet PageSetup | C# transfer printable settings between Excel sheets | clone worksheet printing configuration .NET | reflection copy PageSetup properties C# | copy page margins orientation Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads a workbook, accesses the 'Source' and 'Destination' worksheets, uses reflection to iterate over all writable PageSetup properties, copies each value from the source sheet's PageSetup to the destination sheet's PageSetup, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets["Source"];
            Worksheet destSheet = workbook.Worksheets["Destination"];

            if (sourceSheet == null || destSheet == null)
            {
                Console.WriteLine("Required worksheets 'Source' or 'Destination' are missing.");
                return;
            }

            // Copy all writable PageSetup properties from source to destination
            PageSetup srcSetup = sourceSheet.PageSetup;
            PageSetup dstSetup = destSheet.PageSetup;

            PropertyInfo[] props = typeof(PageSetup).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in props)
            {
                // Skip read‑only or indexed properties
                if (prop.CanRead && prop.CanWrite && prop.GetIndexParameters().Length == 0)
                {
                    object value = prop.GetValue(srcSetup);
                    prop.SetValue(dstSetup, value);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
