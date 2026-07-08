using System;
using Aspose.Cells;

namespace AsposeCellsRichTextSummary
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or iterate through all worksheets as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range to limit iteration
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            Console.WriteLine("Rich Text Summary for Worksheet: " + worksheet.Name);
            Console.WriteLine("---------------------------------------------------");

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Check if the cell contains rich text
                    if (cell.IsRichText())
                    {
                        // Retrieve all character formatting segments
                        FontSetting[] segments = cell.GetCharacters();

                        // Output information for each segment
                        foreach (FontSetting segment in segments)
                        {
                            string textSegment = cell.StringValue.Substring(segment.StartIndex, segment.Length);
                            string fontName = segment.Font.Name;

                            Console.WriteLine($"Cell {cell.Name}:");
                            Console.WriteLine($"  Text Segment : \"{textSegment}\"");
                            Console.WriteLine($"  Start Index  : {segment.StartIndex}");
                            Console.WriteLine($"  Length       : {segment.Length}");
                            Console.WriteLine($"  Font Name    : {fontName}");
                            Console.WriteLine();
                        }
                    }
                }
            }

            // Save the workbook (optional, e.g., after modifications)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}