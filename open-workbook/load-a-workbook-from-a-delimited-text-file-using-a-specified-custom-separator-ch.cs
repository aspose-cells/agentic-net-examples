using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomSeparatorDemo
{
    class Program
    {
        static void Main()
        {
            // Define the custom separator character
            char customSeparator = '|';

            // Prepare sample delimited text data using the custom separator
            string[] lines =
            {
                "Name|Age|City",
                "John Doe|30|New York",
                "Jane Smith|25|London"
            };
            string filePath = "sample_data.txt";

            // Write the sample data to a text file
            File.WriteAllLines(filePath, lines);

            // Create TxtLoadOptions and set the custom separator
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = customSeparator;

            // Load the workbook from the delimited text file using the load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Output some loaded values to verify correct parsing
            Console.WriteLine("A1 (Header): " + cells["A1"].StringValue); // Expected: Name
            Console.WriteLine("B2 (Age): " + cells["B2"].StringValue);   // Expected: 30
            Console.WriteLine("C3 (City): " + cells["C3"].StringValue); // Expected: London

            // Optionally, save the workbook to an Excel file to confirm successful load
            workbook.Save("LoadedFromCustomSeparator.xlsx", SaveFormat.Xlsx);
        }
    }
}