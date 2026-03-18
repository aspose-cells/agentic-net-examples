using System;
using System.IO;
using System.Text;
using Aspose.Cells;

public class CsvImportNoNumericConversion
{
    public static void Run()
    {
        // Sample CSV content containing numbers and strings
        string csvData = "ID,Name,Price,Quantity\n" +
                         "1,Product A,19.99,5\n" +
                         "2,Product B,\"24.50\",\"10\"\n" +
                         "3,Product C,15.75,\"N/A\"";

        // Create a memory stream from the CSV string
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            // Configure load options to disable automatic numeric conversion
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                ConvertNumericData = false,   // Prevent numbers from being converted
                ConvertDateTimeData = false   // Optional: keep dates as strings as well
            };

            // Load the CSV data into a workbook using the specified options
            Workbook workbook = new Workbook(csvStream, loadOptions);

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Verify that values are treated as strings (type should be System.String)
            Console.WriteLine("Cell A2 (ID) type: " + cells["A2"].Value.GetType());
            Console.WriteLine("Cell C2 (Price) type: " + cells["C2"].Value.GetType());
            Console.WriteLine("Cell D3 (Quantity) type: " + cells["D3"].Value.GetType());

            // Save the workbook (optional, demonstrates full lifecycle)
            workbook.Save("NoNumericConversion.xlsx", SaveFormat.Xlsx);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        CsvImportNoNumericConversion.Run();
    }
}