// Title: Generate a summary worksheet that lists each shape’s type, coordinates, size, and adjustment values using Aspose.Cells for .NET
// AI Prompts: Create C# code with Aspose.Cells that scans every worksheet, extracts each shape’s name, type, left, top, width, height, and any adjustment values, and writes the data to a new sheet named "Summary". | Extend the shape‑report program to also retrieve each shape’s rotation angle and add it as an extra column in the summary worksheet. | Add comprehensive error handling and logging to the shape extraction loop to record shapes missing adjustment data or throwing exceptions, and write the log to a text file.
// Common Searches: Aspose.Cells C# how to list all shapes with their positions and sizes in an Excel workbook | C# extract shape adjustment values from Excel using Aspose.Cells | Create a summary sheet of shape properties (type, left, top, width, height) with Aspose.Cells | Iterate through worksheets and shapes in Aspose.Cells to generate a shape report | Reflection to access shape adjustments in Aspose.Cells for .NET
// Tags: Aspose.Cells shape property extraction | C# export shape dimensions to Excel | Aspose.Cells reflection for shape adjustments | Excel shape summary worksheet generation | capture shape rotation Aspose.Cells | robust shape extraction error handling

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing workbook (or creates a new one), adds a "Summary" worksheet, iterates through all other worksheets and their shapes, collects each shape's name, type, left, top, width, height, and any adjustment values via reflection, writes these details into the summary sheet, and saves the workbook.
class ShapeReportGenerator
{
    static void Main()
    {
        try
        {
            // Load the source workbook; create a new one if the file does not exist
            Workbook workbook;
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default sheet
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Add a new worksheet for the summary report
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header row
            summarySheet.Cells["A1"].PutValue("Worksheet");
            summarySheet.Cells["B1"].PutValue("Shape Name");
            summarySheet.Cells["C1"].PutValue("Shape Type");
            summarySheet.Cells["D1"].PutValue("Left");
            summarySheet.Cells["E1"].PutValue("Top");
            summarySheet.Cells["F1"].PutValue("Width");
            summarySheet.Cells["G1"].PutValue("Height");
            summarySheet.Cells["H1"].PutValue("Adjustments");

            int currentRow = 1; // Start writing data from the second row (zero‑based index)

            // Iterate through all worksheets except the summary sheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                if (ws.Name == "Summary")
                    continue; // Skip the summary sheet itself

                // Iterate through all shapes in the current worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    try
                    {
                        // Collect adjustment values (if any) into a comma‑separated string
                        string adjustments = string.Empty;
                        // The Adjustments property may not be available in older versions;
                        // if it exists, gather its values.
                        var adjustmentsProp = shape.GetType().GetProperty("Adjustments");
                        if (adjustmentsProp != null)
                        {
                            var adjustmentsObj = adjustmentsProp.GetValue(shape);
                            if (adjustmentsObj != null)
                            {
                                var countProp = adjustmentsObj.GetType().GetProperty("Count");
                                var indexer = adjustmentsObj.GetType().GetProperty("Item");
                                int count = (int)countProp.GetValue(adjustmentsObj);
                                var sb = new StringBuilder();
                                for (int adjIndex = 0; adjIndex < count; adjIndex++)
                                {
                                    double value = (double)indexer.GetValue(adjustmentsObj, new object[] { adjIndex });
                                    if (sb.Length > 0) sb.Append(", ");
                                    sb.Append(value);
                                }
                                adjustments = sb.ToString();
                            }
                        }

                        // Write shape information to the summary sheet
                        summarySheet.Cells[currentRow, 0].PutValue(ws.Name);                         // Worksheet
                        summarySheet.Cells[currentRow, 1].PutValue(shape.Name);                     // Shape Name
                        summarySheet.Cells[currentRow, 2].PutValue(shape.Type.ToString());          // Shape Type
                        summarySheet.Cells[currentRow, 3].PutValue(shape.Left);                     // Left
                        summarySheet.Cells[currentRow, 4].PutValue(shape.Top);                      // Top
                        summarySheet.Cells[currentRow, 5].PutValue(shape.Width);                    // Width
                        summarySheet.Cells[currentRow, 6].PutValue(shape.Height);                   // Height
                        summarySheet.Cells[currentRow, 7].PutValue(adjustments);                    // Adjustments

                        currentRow++;
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Error processing shape '{shape.Name}' in worksheet '{ws.Name}': {shapeEx.Message}");
                    }
                }
            }

            // Save the workbook with the new summary worksheet
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Summary report saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
