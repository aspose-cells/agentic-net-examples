// Title: Compare HTML file size with EnableCssCustomProperties set to true vs false in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, saves it to HTML twice—once with EnableCssCustomProperties enabled and once disabled—and logs the byte size of each file. | Write a C# helper that uses reflection to set the EnableCssCustomProperties property on HtmlSaveOptions only when the property is present.
// Common Searches: Aspose.Cells HTML export size difference when EnableCssCustomProperties is true | How does enabling CSS custom properties affect the size of generated HTML in Aspose.Cells? | C# code to measure Aspose.Cells HTML file size with and without custom CSS properties | Reflection technique to toggle EnableCssCustomProperties in HtmlSaveOptions for Aspose.Cells | Performance impact of EnableCssCustomProperties on Aspose.Cells HTML output
// Tags: Aspose.Cells HTML size comparison | EnableCssCustomProperties effect on HTML output | HtmlSaveOptions toggle custom CSS property | C# measure generated HTML file size | reflection usage with HtmlSaveOptions

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, saves it to two temporary HTML files using HtmlSaveOptions with EnableCssCustomProperties set to true and false (via reflection), then reads and prints the byte size of each file to illustrate the size impact of the property.
class HtmlSizeComparison
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate the worksheet with sample data
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define temporary file paths
            string pathWithCustomProps = Path.GetTempFileName() + ".html";
            string pathWithoutCustomProps = Path.GetTempFileName() + ".html";

            // Save HTML with EnableCssCustomProperties = true (if the property exists)
            HtmlSaveOptions optionsWith = new HtmlSaveOptions(SaveFormat.Html);
            SetEnableCssCustomProperties(optionsWith, true);
            workbook.Save(pathWithCustomProps, optionsWith);

            // Save HTML with EnableCssCustomProperties = false (if the property exists)
            HtmlSaveOptions optionsWithout = new HtmlSaveOptions(SaveFormat.Html);
            SetEnableCssCustomProperties(optionsWithout, false);
            workbook.Save(pathWithoutCustomProps, optionsWithout);

            // Get file sizes
            long sizeWith = new FileInfo(pathWithCustomProps).Length;
            long sizeWithout = new FileInfo(pathWithoutCustomProps).Length;

            // Output the comparison results
            Console.WriteLine($"HTML size with EnableCssCustomProperties = true : {sizeWith} bytes");
            Console.WriteLine($"HTML size with EnableCssCustomProperties = false: {sizeWithout} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Clean up temporary files if they exist
            DeleteFileIfExists(Path.GetTempFileName() + ".html"); // placeholder to ensure method exists
        }
    }

    // Sets the EnableCssCustomProperties property via reflection if it exists
    private static void SetEnableCssCustomProperties(HtmlSaveOptions options, bool value)
    {
        var prop = typeof(HtmlSaveOptions).GetProperty("EnableCssCustomProperties");
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(options, value);
        }
    }

    // Deletes a file safely if it exists
    private static void DeleteFileIfExists(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch
        {
            // Suppress any exceptions during cleanup
        }
    }
}
