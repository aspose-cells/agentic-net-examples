using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class LoadCsvWithoutNumericConversion
{
    static void Main()
    {
        // Sample CSV data containing numeric-looking strings
        string csvData = "ID,Name,Price,Quantity\n" +
                         "1,Product A,19.99,5\n" +
                         "2,Product B,\"24.50\",\"10\"\n" +
                         "3,Product C,15.75,\"N/A\"";

        // Create TxtLoadOptions for CSV and disable automatic numeric conversion
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        loadOptions.ConvertNumericData = false;   // Prevent conversion of numeric strings
        loadOptions.ConvertDateTimeData = false; // Keep date strings as strings (optional)

        // Load the CSV from a memory stream using the specified options
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(csvStream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Demonstrate that a numeric-looking cell remains a string
            Console.WriteLine("Cell C2 Type : " + worksheet.Cells["C2"].Type);        // Expected: String
            Console.WriteLine("Cell C2 Value: " + worksheet.Cells["C2"].StringValue); // Expected: "19.99"

            // Save the workbook to an Excel file
            workbook.Save("OutputWithoutNumericConversion.xlsx");
        }
    }
}