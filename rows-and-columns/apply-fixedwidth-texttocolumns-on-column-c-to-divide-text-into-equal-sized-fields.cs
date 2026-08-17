// Title: Split Fixed‑Width Text in Column C into Equal‑Sized Fields with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts fixed‑width strings into column C, defines equal field lengths (e.g., 4 characters), loops through each row to extract substrings, writes each segment to adjacent cells, auto‑fits the new columns, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | fixed width | TextToColumns | split column | equal column width | Excel automation | parse fixed‑width data | auto fit columns | workbook save
// Common Searches: Aspose.Cells C# split fixed width column | How to divide a fixed‑width string into multiple Excel columns using Aspose.Cells | Apply equal‑width TextToColumns in Aspose.Cells for .NET | Auto‑fit columns after parsing fixed‑width data with Aspose.Cells | C# example for extracting 4‑character fields from column C
// Developer Intent: The developer needs to break a fixed‑width string in column C into separate, equally sized columns programmatically.
// Use Cases: Transform legacy fixed‑width reports into a structured Excel table for analysis. | Separate concatenated product codes (e.g., ABCD1234) into distinct parts for downstream calculations. | Parse fixed‑length identifiers into individual fields before applying business rules.
// AI Prompts: Generate C# code that uses Aspose.Cells to split a fixed‑width column into three 4‑character columns. | Show an Aspose.Cells example that extracts equal‑size substrings from column C and writes them to adjacent cells. | Provide a .NET snippet that auto‑fits columns after dividing fixed‑width text with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFixedWidthDemo
{
    // Creates a workbook, inserts fixed‑width strings into column C, defines equal field lengths (e.g., 4 characters), loops through each row to extract substrings, writes each segment to adjacent cells, auto‑fits the new columns, and saves the file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data in column C (index 2) with fixed‑width fields (e.g., 4 characters each)
                cells["C1"].PutValue("ABCD1234WXYZ");
                cells["C2"].PutValue("EFGH5678UVWX");
                cells["C3"].PutValue("IJKL9012QRST");

                // Define equal column widths (4 characters each)
                int[] columnWidths = new int[] { 4, 4, 4 };

                // Determine the number of rows to process (include all data rows)
                int totalRows = cells.MaxDataRow + 1;

                // Manually split fixed‑width data from column C into subsequent columns
                for (int row = 0; row < totalRows; row++)
                {
                    string source = cells[row, 2].StringValue ?? string.Empty;
                    for (int i = 0; i < columnWidths.Length; i++)
                    {
                        int start = i * columnWidths[i];
                        if (start >= source.Length)
                            break;

                        int length = Math.Min(columnWidths[i], source.Length - start);
                        string part = source.Substring(start, length);
                        cells[row, 2 + i].PutValue(part);
                    }
                }

                // Optional: autofit the newly created columns for better visibility
                sheet.AutoFitColumns(2, columnWidths.Length); // columns C onward

                // Prepare output path and ensure the directory exists
                string outputPath = "FixedWidthOutput.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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
