// Title: C# – Apply a Yellow Bold Style to /Customer/Name Cells Mapped via XML in Aspose.Cells
// Description: Loads an existing workbook, creates a solid‑yellow background with bold font style, defines the range that corresponds to the /Customer/Name XML element, applies the style to every cell in that range, and saves the workbook while handling missing‑file errors.
// Keywords: Aspose.Cells C# XML mapping | apply custom style to Excel cells | highlight customer name column | yellow background bold font Aspose.Cells | range styling based on XML map | C# Excel formatting Aspose.Cells
// Common Searches: Aspose.Cells style cells mapped to XML element | C# highlight /Customer/Name column in Excel | apply yellow background to Excel range using Aspose.Cells | how to format XML‑mapped cells in .NET | Aspose.Cells custom style example C#
// Developer Intent: Highlight the cells that represent the /Customer/Name element by applying a yellow background and bold font using Aspose.Cells for .NET.
// Use Cases: Make customer names stand out in reports generated from XML data. | Visually separate name fields from other mapped data after importing XML into Excel. | Ensure consistent formatting across multiple workbooks processed programmatically.
// AI Prompts: Generate a reusable C# method in Aspose.Cells that applies a yellow‑bold style to any range identified by an XML map path. | Show how to retrieve the exact cell range for /Customer/Name dynamically instead of using a hard‑coded address. | Provide best‑practice error handling for loading workbooks and applying styles with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, creates a solid‑yellow background with bold font style, defines the range that corresponds to the /Customer/Name XML element, applies the style to every cell in that range, and saves the workbook while handling missing‑file errors.
class ApplyCustomStyleToCustomerNames
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains the mapped XML data
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a custom style (yellow background, bold font) to highlight names
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Font.IsBold = true;

            // Define the cell range that corresponds to the /Customer/Name element.
            // Example assumes customer names are placed in column B, rows 2 through 100.
            Aspose.Cells.Range nameRange = cells.CreateRange("B2:B100");

            // Apply the custom style to all cells in the defined range
            StyleFlag flag = new StyleFlag { All = true };
            nameRange.ApplyStyle(highlightStyle, flag);

            // Save the workbook with the applied styling
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
