// Title: C# – Append the noadd parameter to the first smart‑marker template row to keep the header static with Aspose.Cells
// Common Searches: how to add the noadd flag to a smart marker template row in Aspose.Cells C# | keep header row static during smart marker merge using Aspose.Cells .NET | append noadd parameter to cell A2 smart marker with Aspose.Cells library | prevent row addition when processing smart markers in an Excel workbook C# | modify smart marker text to include noadd in Aspose.Cells example
// Tags: smart marker noadd option Aspose.Cells | static header during smart marker merge .NET | modify template cell A2 Aspose.Cells | disable row insertion with smart markers | update smart marker text C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, reads the smart marker in cell A2, appends the ",noadd" flag if it is missing to keep the header row from being duplicated during merging, and saves the modified file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                // Handle password‑protected or other Cells specific errors
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the template cell that contains a smart marker
                Cell templateCell = worksheet.Cells["A2"];

                // Retrieve the existing smart marker text
                string markerText = templateCell.StringValue;

                // Append ",noadd" if not already present
                if (!string.IsNullOrEmpty(markerText) && !markerText.Contains("noadd"))
                {
                    string updatedMarker = markerText.Contains(",")
                        ? $"{markerText},noadd"
                        : $"{markerText},noadd";

                    templateCell.PutValue(updatedMarker);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
