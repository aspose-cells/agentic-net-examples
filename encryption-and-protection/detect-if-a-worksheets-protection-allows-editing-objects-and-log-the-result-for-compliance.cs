// Title: Check if an Excel worksheet allows editing objects using Aspose.Cells for .NET and log the result
// AI Prompts: Write a C# program with Aspose.Cells that opens a workbook, uses reflection to read the Protection.AllowEditObject flag of the first worksheet, and outputs a compliance message. | Create a .NET snippet that loads an .xlsx file, determines whether object editing is permitted on a worksheet via the Protection API, prints the status, and saves the file.
// Common Searches: Aspose.Cells how to determine if worksheet protection allows object editing in C# | C# read AllowEditObject flag from Excel sheet protection using Aspose | Check worksheet edit objects permission with Aspose.Cells .NET | Log worksheet protection settings for object editing in Aspose.Cells | Reflection get AllowEditObject property Aspose.Cells workbook
// Tags: Aspose.Cells worksheet protection AllowEditObject | C# reflection read Protection property | detect edit objects permission Excel Aspose | log worksheet protection compliance .NET | save workbook after protection check Aspose

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx with Aspose.Cells, uses reflection to read the Protection.AllowEditObject flag of the first worksheet, prints whether editing objects is allowed, and saves the workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine if the worksheet's protection permits editing objects
        bool allowsEditObjects = false;
        try
        {
            var prop = typeof(Protection).GetProperty("AllowEditObject");
            if (prop != null && prop.PropertyType == typeof(bool))
            {
                allowsEditObjects = (bool)prop.GetValue(sheet.Protection);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving protection property: {ex.Message}");
        }

        // Log the compliance result
        if (allowsEditObjects)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' allows editing objects.");
        }
        else
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' does NOT allow editing objects.");
        }

        // Save the workbook (no modifications made, but required by lifecycle rules)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
