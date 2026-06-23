using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceDraftInNamedRange
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range
                Name publishRange = workbook.Worksheets.Names["PublishRange"];
                if (publishRange == null)
                {
                    Console.WriteLine("Named range 'PublishRange' not found.");
                    return;
                }

                // Get reference in A1 style and remove leading '='
                string refersTo = publishRange.GetRefersTo(false, false);
                if (refersTo.StartsWith("="))
                    refersTo = refersTo.Substring(1);

                // Split into sheet name and address
                string[] parts = refersTo.Split('!');
                if (parts.Length != 2)
                {
                    Console.WriteLine("Invalid named range reference.");
                    return;
                }

                string sheetName = parts[0];
                string rangeAddress = parts[1];

                // Access the worksheet
                Worksheet ws = workbook.Worksheets[sheetName];
                if (ws == null)
                {
                    Console.WriteLine($"Worksheet '{sheetName}' not found.");
                    return;
                }

                // Create Aspose.Cells.Range explicitly
                Aspose.Cells.Range range = ws.Cells.CreateRange(rangeAddress);

                // Replace "Draft" with "Final" in string cells
                foreach (Cell cell in range)
                {
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        if (original.Contains("Draft"))
                        {
                            cell.PutValue(original.Replace("Draft", "Final"));
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}