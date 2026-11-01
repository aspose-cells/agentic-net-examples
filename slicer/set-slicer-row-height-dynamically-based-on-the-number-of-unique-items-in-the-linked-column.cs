// Title: Resize an Excel slicer height dynamically based on unique values in a linked column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set slicer height from distinct column entries | How to programmatically adjust Excel slicer size according to unique items | Calculate slicer row height in .NET by counting unique values | Automatically resize slicer height after retrieving unique data with Aspose.Cells | C# example for dynamic slicer height based on linked column count
// Tags: slicer height adjustment Aspose.Cells | unique column value count C# | dynamic slicer resizing .NET | excel slicer row height calculation | Aspose.Cells slicer size based on data

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerResize
{
    // The code loads an Excel workbook, identifies the worksheet linked to a slicer, counts distinct values in a specified column, multiplies that count by a predefined row height, and sets the slicer's Height property accordingly before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";
            const string slicerName = "Slicer1"; // replace with actual slicer name
            const int linkedColumnIndex = 0; // column A

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheet that contains the data linked to the slicer
                Worksheet dataSheet = workbook.Worksheets["Data"];
                if (dataSheet == null)
                {
                    Console.WriteLine("Worksheet 'Data' not found.");
                    return;
                }

                // Collect unique values from the linked column (skip header row)
                int firstDataRow = dataSheet.Cells.MaxDataRow > 0 ? 1 : 0; // assume header at row 0
                int lastDataRow = dataSheet.Cells.MaxDataRow;
                HashSet<object> uniqueValues = new HashSet<object>();

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    object val = dataSheet.Cells[row, linkedColumnIndex].Value;
                    if (val != null)
                        uniqueValues.Add(val);
                }

                int uniqueCount = uniqueValues.Count;

                // Desired height for a single slicer row (in points)
                double singleRowHeight = 20.0; // adjust as needed

                // Retrieve the slicer by name
                Slicer slicer = dataSheet.Slicers[slicerName];
                if (slicer == null)
                {
                    Console.WriteLine($"Slicer '{slicerName}' not found on worksheet '{dataSheet.Name}'.");
                    return;
                }

                // Set the slicer's overall height based on the number of unique items
                slicer.Height = singleRowHeight * uniqueCount;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
