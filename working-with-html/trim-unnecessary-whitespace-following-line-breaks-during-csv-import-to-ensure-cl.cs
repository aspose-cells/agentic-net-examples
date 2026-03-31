using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvTrimExample
{
    class Program
    {
        static void Main()
        {
            // Sample CSV data containing spaces after line breaks
            string csvData = "Name, Age, City\r\n" +
                             "John, 30, New York\r\n" +
                             "Alice, 25,   London\r\n" +
                             "Bob, 35,Paris   \r\n";

            // Convert CSV string to a memory stream
            using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Configure load options for CSV import
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    Separator = ',',               // Use comma as delimiter
                    HasTextQualifier = true,       // Enable text qualifier handling
                    TreatConsecutiveDelimitersAsOne = false
                };

                // Import CSV data into the first worksheet starting at cell A1
                workbook.Worksheets[0].Cells.ImportCSV(csvStream, loadOptions, 0, 0);

                // Trim leading and trailing whitespace from all string cells
                Cells cells = workbook.Worksheets[0].Cells;
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string trimmed = cell.StringValue.Trim();
                            if (trimmed != cell.StringValue)
                            {
                                cell.PutValue(trimmed);
                            }
                        }
                    }
                }

                // Save the cleaned workbook to an XLSX file
                workbook.Save("CleanedData.xlsx", SaveFormat.Xlsx);
            }

            Console.WriteLine("CSV imported and whitespace trimmed successfully.");
        }
    }
}