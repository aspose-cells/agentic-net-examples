using System;
using Aspose.Cells;

namespace AsposeCellsPlainTextDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the plain text (CSV) file
            string txtFilePath = "sample.txt";

            // Create load options for a text file
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            // Define the column separator (comma for CSV)
            loadOptions.Separator = ',';

            // Open the workbook from the plain text file using the load options
            Workbook workbook = new Workbook(txtFilePath, loadOptions);

            // Display basic information to verify the load succeeded
            Console.WriteLine("Worksheets count: " + workbook.Worksheets.Count);
            Console.WriteLine("Value of cell A1: " + workbook.Worksheets[0].Cells["A1"].StringValue);

            // Save the loaded workbook to an Excel file
            workbook.Save("ConvertedFromText.xlsx", SaveFormat.Xlsx);
        }
    }
}