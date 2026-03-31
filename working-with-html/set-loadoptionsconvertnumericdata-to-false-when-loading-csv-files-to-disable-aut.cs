using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample CSV content with numeric-like strings
        string csvData = "ID,Price,Quantity\n" +
                         "1,19.99,5\n" +
                         "2,24.50,10\n" +
                         "3,N/A,";

        // Convert the CSV string to a memory stream
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            // Create TxtLoadOptions for CSV files
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            // Disable automatic conversion of string values to numeric data
            loadOptions.ConvertNumericData = false;

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(csvStream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Demonstrate that values remain strings
            Console.WriteLine($"Cell A2 (ID)   : {worksheet.Cells["A2"].StringValue} (Type: {worksheet.Cells["A2"].Type})");
            Console.WriteLine($"Cell B2 (Price): {worksheet.Cells["B2"].StringValue} (Type: {worksheet.Cells["B2"].Type})");
            Console.WriteLine($"Cell C2 (Qty)  : {worksheet.Cells["C2"].StringValue} (Type: {worksheet.Cells["C2"].Type})");

            // Save the workbook if needed
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}