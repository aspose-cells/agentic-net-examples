// Title: Read slicer properties, log details, and adjust row height from slicer position using Aspose.Cells for .NET (C#)
// Common Searches: how to enumerate slicers and retrieve their properties with Aspose.Cells in C# | set Excel row height based on slicer location using Aspose.Cells .NET | log slicer name, caption, size and linked table/column with Aspose.Cells | use reflection to access TableName and ColumnName of a slicer in Aspose.Cells | adjust worksheet row height programmatically after reading slicer dimensions
// Tags: Aspose.Cells read slicer metadata C# | Aspose.Cells adjust row height from slicer position | Aspose.Cells reflection access slicer TableName ColumnName | Aspose.Cells log slicer name caption style | Aspose.Cells modify worksheet based on slicer dimensions

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The example loads an Excel workbook, iterates through all slicers on the first worksheet, and logs each slicer's Name, Caption, Height, Width, linked TableName, ColumnName, and Style (using reflection for version‑safe access). It then obtains the slicer's UpperLeftRow and sets that row's height to half the slicer's height before saving the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of slicers on the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate through each slicer
            foreach (Slicer slicer in slicers)
            {
                try
                {
                    // Log basic slicer properties
                    Console.WriteLine($"Slicer Name: {slicer.Name}");
                    Console.WriteLine($"Slicer Caption: {slicer.Caption}");
                    Console.WriteLine($"Slicer Size - Height: {slicer.Height}, Width: {slicer.Width}");

                    // The following properties may not be available in older Aspose.Cells versions.
                    // They are accessed conditionally using reflection to avoid compile‑time errors.

                    // Table name
                    var tableNameProp = slicer.GetType().GetProperty("TableName");
                    if (tableNameProp != null)
                    {
                        var tableName = tableNameProp.GetValue(slicer) as string;
                        Console.WriteLine($"Linked Table Name: {tableName}");
                    }

                    // Column name
                    var columnNameProp = slicer.GetType().GetProperty("ColumnName");
                    if (columnNameProp != null)
                    {
                        var columnName = columnNameProp.GetValue(slicer) as string;
                        Console.WriteLine($"Linked Column Name: {columnName}");
                    }

                    // Slicer style
                    var styleProp = slicer.GetType().GetProperty("SlicerStyle");
                    if (styleProp != null)
                    {
                        var style = styleProp.GetValue(slicer);
                        Console.WriteLine($"Slicer Style: {style}");
                    }

                    // Position (row/column) – also accessed via reflection
                    var rowProp = slicer.GetType().GetProperty("UpperLeftRow");
                    var colProp = slicer.GetType().GetProperty("UpperLeftColumn");
                    if (rowProp != null && colProp != null)
                    {
                        int row = (int)rowProp.GetValue(slicer);
                        int col = (int)colProp.GetValue(slicer);
                        Console.WriteLine($"Slicer Position - Row: {row}, Column: {col}");

                        // Example logic: set the height of the row where the slicer starts
                        double newRowHeight = slicer.Height * 0.5;
                        worksheet.Cells.SetRowHeight(row, newRowHeight);
                        Console.WriteLine($"Row {row} height set to {newRowHeight}");
                    }

                    Console.WriteLine(new string('-', 40));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing slicer '{slicer?.Name}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
