// Title: Convert a pipe‑delimited TXT workbook to an Excel XML spreadsheet by splitting cells with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells split pipe delimited text file into separate columns and save as XML | How to load a TXT workbook with a custom delimiter using Aspose.Cells .NET | Aspose.Cells parse delimited cell values and export to Excel XML format | Convert pipe‑separated text to an Excel XML spreadsheet in C# | Split cell contents by '|' while loading a TXT file with Aspose.Cells and export to XML
// Tags: pipe delimited txt to excel xml conversion Aspose.Cells | split cell values by custom delimiter Aspose.Cells | load txt workbook and save as xml Aspose.Cells | cell splitting and column insertion Aspose.Cells | excel xml spreadsheet export using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads a pipe‑delimited TXT file as a workbook with Aspose.Cells, iterates through each used cell, splits values containing the '|' delimiter into separate parts, places the first part back into the original cell and the remaining parts into adjacent columns, and finally saves the modified workbook as an Excel XML Spreadsheet.
class TxtToXmlConverter
{
    static void Main()
    {
        // Path to the source TXT file
        string txtPath = "input.txt";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(txtPath))
        {
            Console.WriteLine($"Input file not found: {txtPath}");
            return;
        }

        // Define the custom delimiter used in the TXT file (e.g., '|')
        char delimiter = '|';

        try
        {
            // Load the TXT file as a workbook. Aspose.Cells can auto‑detect the format.
            Workbook workbook = new Workbook(txtPath);

            // Get the first worksheet (the TXT data is loaded into the first sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Iterate through each cell in the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    string cellValue = cell.StringValue;

                    // If the cell contains the delimiter, split it into parts
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(delimiter.ToString()))
                    {
                        string[] parts = cellValue.Split(delimiter);

                        // Replace the original cell with the first part
                        cell.PutValue(parts[0]);

                        // Insert the remaining parts into subsequent columns
                        for (int i = 1; i < parts.Length; i++)
                        {
                            int targetCol = col + i;
                            sheet.Cells[row, targetCol].PutValue(parts[i]);
                        }
                    }
                }
            }

            // Save the processed workbook as an XML file (Excel XML Spreadsheet format)
            string xmlPath = "output.xml";
            workbook.Save(xmlPath, SaveFormat.Xml);

            Console.WriteLine("Conversion completed. XML saved to: " + xmlPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during conversion: " + ex.Message);
        }
    }
}
