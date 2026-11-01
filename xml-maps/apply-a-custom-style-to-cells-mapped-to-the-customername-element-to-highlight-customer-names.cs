// Title: Apply a yellow bold highlight to Excel cells mapped to the /Customer/Name XML element using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, uses Aspose.Cells to locate cells mapped to the '/Customer/Name' XML path, creates a yellow solid background with bold font style, and applies it to those cells. | Write a C# snippet that accesses the XmlMaps collection via reflection in Aspose.Cells, retrieves the CellArea array for a specific XML element, and sets a custom style on each cell in the area.
// Common Searches: how to use Aspose.Cells to style cells mapped to an XML element in C# | apply background color to XML mapped cells in an Excel file using Aspose.Cells | retrieve cell areas for /Customer/Name XML map with Aspose.Cells .NET | highlight customer name fields in Excel via XML map using Aspose.Cells | C# reflection to access XmlMaps collection in Aspose.Cells
// Tags: apply style to XML mapped cells Aspose.Cells | highlight customer name cells Excel Aspose.Cells | retrieve mapped cell areas XmlMap C# | reflection access XmlMaps Aspose.Cells | custom yellow background style Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook, uses reflection to obtain its XmlMaps collection, fetches the cell areas mapped to the '/Customer/Name' XML element, creates a yellow solid background with bold font style, applies this style to each mapped cell, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to access XML maps via reflection (compatible with all Aspose.Cells versions)
            PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp != null)
            {
                object xmlMapsObj = xmlMapsProp.GetValue(workbook);
                if (xmlMapsObj != null)
                {
                    // Get the Count property of the XmlMapCollection
                    PropertyInfo countProp = xmlMapsObj.GetType().GetProperty("Count");
                    int mapCount = countProp != null ? (int)countProp.GetValue(xmlMapsObj) : 0;

                    if (mapCount > 0)
                    {
                        // Retrieve the first XmlMap (indexer property "Item")
                        PropertyInfo itemProp = xmlMapsObj.GetType().GetProperty("Item");
                        object xmlMapObj = itemProp?.GetValue(xmlMapsObj, new object[] { 0 });

                        if (xmlMapObj != null)
                        {
                            dynamic xmlMap = xmlMapObj; // Use dynamic to call members

                            // Retrieve all cell areas mapped to the specified XML element
                            CellArea[] mappedAreas = xmlMap.GetMappedCells("/Customer/Name");

                            // Get the worksheet associated with this XML map
                            Worksheet sheet = xmlMap.Worksheet;

                            // Create a style to highlight the mapped cells
                            Style highlightStyle = workbook.CreateStyle();
                            highlightStyle.ForegroundColor = Color.Yellow;          // Background color
                            highlightStyle.Pattern = BackgroundType.Solid;          // Apply solid fill
                            highlightStyle.Font.IsBold = true;                      // Bold font

                            // Apply the style to each cell within the mapped areas
                            foreach (CellArea area in mappedAreas)
                            {
                                for (int row = area.StartRow; row <= area.EndRow; row++)
                                {
                                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                                    {
                                        Cell cell = sheet.Cells[row, col];
                                        cell.SetStyle(highlightStyle);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No XML maps found in the workbook.");
                    }
                }
                else
                {
                    Console.WriteLine("Workbook does not contain XML maps.");
                }
            }
            else
            {
                Console.WriteLine("The loaded Aspose.Cells version does not support XML maps.");
            }

            // Ensure the output directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
