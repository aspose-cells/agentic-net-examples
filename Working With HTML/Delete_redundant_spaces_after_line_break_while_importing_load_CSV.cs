using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteRedundantSpacesAfterLineBreakCsv
    {
        public static void Run()
        {
            // Sample CSV data with redundant spaces after line breaks
            string csvData = "Name, Age, City\nJohn,  30,   New York\n  Alice,25,   London\nBob,   35,Paris";
            byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(csvData);
            using (MemoryStream csvStream = new MemoryStream(csvBytes))
            {
                // Set up TxtLoadOptions for CSV import
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
                loadOptions.Separator = ',';               // Use comma as delimiter
                loadOptions.ConvertNumericData = true;     // Convert numeric strings to numbers

                // Load CSV data into a new workbook
                Workbook workbook = new Workbook(csvStream, loadOptions);

                // Iterate through all used cells and collapse multiple spaces into a single space
                Cells cells = workbook.Worksheets[0].Cells;
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            // Replace sequences of two or more spaces with a single space and trim ends
                            string cleaned = Regex.Replace(original, @"\s{2,}", " ").Trim();
                            if (cleaned != original)
                            {
                                cell.PutValue(cleaned);
                            }
                        }
                    }
                }

                // Save the processed workbook to an XLSX file
                workbook.Save("ProcessedCsv.xlsx", SaveFormat.Xlsx);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            DeleteRedundantSpacesAfterLineBreakCsv.Run();
        }
    }
}