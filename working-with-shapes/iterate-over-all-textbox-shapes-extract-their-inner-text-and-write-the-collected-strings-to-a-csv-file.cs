// Title: Extract all TextBox shape text from an Excel workbook and save it to a CSV file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds shapes with a non‑empty Text property, and writes each text value as a quoted field to an output CSV file. | Adjust the shape‑iteration example to skip shapes without text and ensure double quotes are escaped correctly when generating CSV rows. | Add comprehensive error handling for workbook loading and CSV writing, including file‑existence verification and detailed exception messages.
// Common Searches: how to export textbox contents from an Excel file to csv using Aspose.Cells C# | iterate through shapes in a workbook and get text property with Aspose.Cells .NET | save extracted shape text from Excel to a CSV file with proper quoting Aspose.Cells
// Tags: Aspose.Cells extract shape text to CSV | C# iterate worksheet shapes Aspose.Cells | export TextBox contents Excel Aspose.Cells | CSV generation with escaped quotes C# | handle empty shape text Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook with Aspose.Cells, loops through all worksheets and their shapes, collects the Text property of each shape that contains text (e.g., TextBox), and writes the collected strings to a CSV file, quoting each field and escaping internal double quotes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputCsv = "output.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect texts from shapes that contain text (e.g., TextBox)
            List<string> extractedTexts = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Add text if the shape has non‑empty Text property
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        extractedTexts.Add(shape.Text);
                    }
                }
            }

            // Write the collected strings to a CSV file (one entry per line)
            try
            {
                using (StreamWriter writer = new StreamWriter(outputCsv))
                {
                    foreach (string line in extractedTexts)
                    {
                        // Escape double quotes and wrap the field in quotes
                        string escaped = $"\"{line.Replace("\"", "\"\"")}\"";
                        writer.WriteLine(escaped);
                    }
                }

                Console.WriteLine($"Extraction completed. Results saved to \"{outputCsv}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing CSV file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
