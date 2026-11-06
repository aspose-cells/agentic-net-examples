// Title: Format a date cell as full month name and day while using the 1904 date system with Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells set custom date format "mmmm d" for a cell | enable 1904 date system in an Aspose.Cells workbook using reflection | apply full month name and day number format to Excel date in C# | save workbook as .xlsx with custom date formatting in Aspose.Cells | how to display March 15 as "March 15" using Aspose.Cells | Aspose.Cells C# example for 1904 date system and custom number format
// Tags: custom number format mmmm d Aspose.Cells | enable 1904 date system via reflection C# | apply style to date cell Aspose.Cells | save workbook with custom date format .xlsx | Aspose.Cells date formatting full month name | C# set workbook 1904 date system

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, uses reflection to enable the 1904 date system when the property is available, writes a DateTime value to cell A1, defines a Style with the custom number format "mmmm d" to show the full month name and day, applies the style to the cell, ensures the output directory exists, and saves the file as DateFormatted1904.xlsx, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable the 1904 date system if the property is available in the current Aspose.Cells version.
            // In some versions the property may be absent; this block safely attempts to set it via reflection.
            try
            {
                var settings = workbook.Settings;
                var prop = settings.GetType().GetProperty("Is1904DateSystem");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(settings, true);
                }
            }
            catch
            {
                // If reflection fails, continue without setting the 1904 date system.
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a sample date value (e.g., 15th March 2023)
            Cell dateCell = sheet.Cells["A1"];
            dateCell.PutValue(new DateTime(2023, 3, 15));

            // Create a style with a custom number format that shows full month name and day (e.g., "March 15")
            Style style = workbook.CreateStyle();
            style.Custom = "mmmm d";

            // Apply the style to the cell
            StyleFlag flag = new StyleFlag { All = true };
            dateCell.SetStyle(style, flag);

            // Define output file path
            string outputPath = "DateFormatted1904.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
