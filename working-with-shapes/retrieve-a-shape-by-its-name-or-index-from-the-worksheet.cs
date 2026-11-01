// Title: How to retrieve an Excel worksheet shape by name or index with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape by its name from worksheet | retrieve shape at specific index using Aspose.Cells .NET | C# code to access worksheet Shapes collection by name | find first shape in Excel file with Aspose.Cells API | Aspose.Cells example for shape name lookup and index lookup
// Tags: retrieve worksheet shape by name Aspose.Cells | access shape by index Aspose.Cells | Aspose.Cells Shapes collection lookup | C# Aspose.Cells shape properties extraction | Excel shape retrieval using Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an existing Excel workbook (or creates a new one), selects the first worksheet, and demonstrates two methods to obtain a Shape object: using the shape's assigned name via the Shapes collection indexer and using its zero‑based index in the collection. It prints each shape's Name and Type when found and saves the workbook to an output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Ensure there is at least one worksheet.
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add("Sheet1");
            }

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // ----- Retrieve shape by its name -----
            // Replace "MyShapeName" with the actual name of the shape you want to get.
            Shape shapeByName = worksheet.Shapes["MyShapeName"];

            // ----- Retrieve shape by its index -----
            // Index is zero‑based. This gets the first shape in the collection.
            Shape shapeByIndex = worksheet.Shapes.Count > 0 ? worksheet.Shapes[0] : null;

            // Display some properties of the retrieved shapes.
            if (shapeByName != null)
            {
                Console.WriteLine($"Shape by name: Name = {shapeByName.Name}, Type = {shapeByName.Type}");
            }
            else
            {
                Console.WriteLine("Shape with the specified name was not found.");
            }

            if (shapeByIndex != null)
            {
                Console.WriteLine($"Shape by index: Name = {shapeByIndex.Name}, Type = {shapeByIndex.Type}");
            }
            else
            {
                Console.WriteLine("No shapes exist at the specified index.");
            }

            // Save the workbook (optional if modifications were made).
            try
            {
                // Ensure the directory for the output file exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
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
