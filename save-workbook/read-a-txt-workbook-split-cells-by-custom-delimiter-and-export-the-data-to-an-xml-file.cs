// Title: Convert a pipe‑delimited TXT workbook to SpreadsheetML XML while splitting cells using Aspose.Cells for .NET
// AI Prompts: Load a pipe‑delimited TXT file into an Aspose.Cells Workbook, split every string cell on the '|' character, and write each part into separate columns of a new workbook. | Traverse all worksheets and rows, copying non‑delimited values unchanged and expanding delimited values into adjacent cells. | Save the populated workbook as SpreadsheetML XML to a target path, creating the output directory if it does not exist.
// Common Searches: Aspose.Cells C# split pipe‑delimited cells when converting TXT to XML | How to export a TXT workbook to SpreadsheetML XML with custom delimiter handling | C# read multi‑sheet TXT file, split cell values by '|' and save as XML using Aspose.Cells | Convert pipe‑separated text workbook to XML preserving original data in Aspose.Cells | Split cell contents during TXT to XML conversion in .NET with Aspose.Cells
// Tags: pipe delimiter cell split Aspose.Cells | TXT to SpreadsheetML XML conversion C# | split cell values during workbook import | multi‑sheet processing Aspose.Cells | preserve non‑delimited cell data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads a pipe‑delimited TXT file into an Aspose.Cells Workbook, iterates through each worksheet and cell, splits string cells on the '|' character into separate columns while copying other values unchanged, and saves the result as a SpreadsheetML XML file.
class TxtToXmlConverter
{
    static void Main()
    {
        // Input TXT file path
        string txtFilePath = @"C:\Data\input.txt";

        // Output XML file path
        string xmlOutputPath = @"C:\Data\output.xml";

        // Custom delimiter used to split cell contents
        string delimiter = "|";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(txtFilePath))
                throw new FileNotFoundException($"Input file not found: {txtFilePath}");

            // Load the TXT workbook (Aspose.Cells automatically detects the format)
            Workbook srcWorkbook = new Workbook(txtFilePath);

            // Create a new workbook that will hold the split data
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            int destRow = 0; // Row index in the destination sheet

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
            {
                // Determine the used range of rows and columns
                int maxRow = srcSheet.Cells.MaxDataRow;
                int maxCol = srcSheet.Cells.MaxDataColumn;

                // Process each row
                for (int i = 0; i <= maxRow; i++)
                {
                    int destCol = 0; // Column index in the destination sheet for the current row

                    // Process each cell in the current row
                    for (int j = 0; j <= maxCol; j++)
                    {
                        Cell srcCell = srcSheet.Cells[i, j];

                        // If the cell contains a string, split it by the custom delimiter
                        if (srcCell.Type == CellValueType.IsString && srcCell.StringValue.Contains(delimiter))
                        {
                            string[] parts = srcCell.StringValue.Split(new string[] { delimiter }, StringSplitOptions.None);
                            foreach (string part in parts)
                            {
                                destSheet.Cells[destRow, destCol].PutValue(part);
                                destCol++;
                            }
                        }
                        else
                        {
                            // Copy the original cell value as‑is
                            destSheet.Cells[destRow, destCol].PutValue(srcCell.Value);
                            destCol++;
                        }
                    }

                    destRow++; // Move to the next row in the destination sheet
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(xmlOutputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the resulting workbook as an XML file (SpreadsheetML format)
            destWorkbook.Save(xmlOutputPath, SaveFormat.Xml);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
