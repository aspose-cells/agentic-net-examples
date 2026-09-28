// Title: Apply the Light2 theme background color to all cells in a specific data table using Aspose.Cells for .NET
// AI Prompts: Create a C# routine that builds a Style with BackgroundColor = Color.LightGray, Pattern = BackgroundType.Solid, and applies it to the range A1:D10 using StyleFlag.All = true in Aspose.Cells. | Write code that loads a workbook, selects a worksheet, defines a range for a data table, applies a LightGray background style to every cell in that range, and saves the updated file.
// Common Searches: Aspose.Cells C# how to set Light2 theme background for a selected range | C# apply solid background style to an Excel table using Aspose.Cells | Using StyleFlag.All to format entire data table in Aspose.Cells .NET | Change background color of cells A1:D10 with Aspose.Cells for .NET | Apply theme color to worksheet range programmatically with Aspose.Cells
// Tags: light2 theme color Aspose.Cells | styleflag all attributes C# | format data table range Aspose.Cells | solid background style Excel C# | apply style to range Aspose.Cells | create workbook style Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // Loads an existing workbook, creates a solid LightGray style, applies it to the A1:D10 data table range with a StyleFlag that enables all style attributes, and saves the modified workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify by name/index)
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the data table range you want to format (adjust as needed)
                AsposeRange dataTable = worksheet.Cells.CreateRange("A1:D10");

                // Create a new style
                Style themeStyle = workbook.CreateStyle();

                // Apply a solid background color (using a light theme-like color)
                themeStyle.BackgroundColor = Color.LightGray;
                themeStyle.Pattern = BackgroundType.Solid;

                // Prepare a flag to apply all style attributes
                StyleFlag flag = new StyleFlag { All = true };

                // Apply the style to the entire data table range
                dataTable.ApplyStyle(themeStyle, flag);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
