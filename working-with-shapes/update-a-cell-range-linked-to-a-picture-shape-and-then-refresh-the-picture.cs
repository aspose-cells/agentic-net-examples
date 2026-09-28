// Title: How to update cells linked to a picture shape and refresh the picture in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that reads a picture's LinkToCell range, iterates through the referenced cells, writes new values, and calls the picture's Refresh method with Aspose.Cells. | Create a C# console program that loads an .xlsx file, finds the first picture shape, parses its linked cell range, updates each cell with a calculated value, and refreshes the picture using dynamic invocation in Aspose.Cells.
// Common Searches: Aspose.Cells C# update linked cell range of a picture and refresh image | Refresh Excel picture after modifying linked cells using Aspose.Cells | How to use LinkToCell property of a picture shape in Aspose.Cells .NET | C# example for parsing picture LinkToCell range and updating cells with Aspose.Cells | Dynamic invocation of Refresh method on picture shape Aspose.Cells
// Tags: Aspose.Cells picture linked cells | invoke picture Refresh method dynamically | C# update cells from picture LinkToCell | parse cell area from string Aspose.Cells | save workbook after picture refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, locates the first picture shape, retrieves its optional LinkToCell range, converts the range string to a CellArea, updates each cell in that area with a computed value, attempts to refresh the picture via dynamic method invocation, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the first picture on the sheet
            Picture? picture = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape is Picture pic)
                {
                    picture = pic;
                    break;
                }
            }

            if (picture != null)
            {
                // Use dynamic to access members that may not exist in older versions
                dynamic dynPic = picture;
                string? linkedRange = null;

                try
                {
                    linkedRange = dynPic.LinkToCell as string;
                }
                catch
                {
                    // Property not available; ignore linking logic
                }

                if (!string.IsNullOrEmpty(linkedRange))
                {
                    // Convert the linked cell range string (e.g., "A1:B2") to a CellArea object
                    CellArea linkedArea;
                    try
                    {
                        string[] parts = linkedRange.Split(':');
                        if (parts.Length == 2)
                        {
                            linkedArea = CellArea.CreateCellArea(parts[0], parts[1]);
                        }
                        else
                        {
                            // Single cell reference
                            linkedArea = CellArea.CreateCellArea(linkedRange, linkedRange);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to parse linked range '{linkedRange}': {ex.Message}");
                        return;
                    }

                    // Update each cell in the linked range
                    for (int row = linkedArea.StartRow; row <= linkedArea.EndRow; row++)
                    {
                        for (int col = linkedArea.StartColumn; col <= linkedArea.EndColumn; col++)
                        {
                            // Example update: set the cell value to row * 10 + column index
                            sheet.Cells[row, col].PutValue(row * 10 + col);
                        }
                    }

                    // Refresh the picture so it reflects the updated cell values
                    try
                    {
                        dynPic.Refresh();
                    }
                    catch
                    {
                        // Refresh method not available; continue without it
                    }
                }
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
