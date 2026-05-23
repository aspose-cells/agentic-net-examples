using System;
using Aspose.Cells;

namespace PipeDelimitedToXlsx
{
    class Program
    {
        static void Main()
        {
            // Path to the pipe‑delimited source file
            string sourcePath = "input_pipe.txt";

            // Configure load options to use pipe character as the separator
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = '|';

            // Load the text file into a workbook using the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the loaded workbook as an XLSX file
            string destinationPath = "output.xlsx";
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"File converted successfully: {destinationPath}");
        }
    }
}