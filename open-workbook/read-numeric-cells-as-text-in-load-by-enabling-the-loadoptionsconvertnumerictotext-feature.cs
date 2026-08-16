// Title: Read CSV numeric cells as text in C# with Aspose.Cells TxtLoadOptions ConvertNumericData = false
// Description: Shows how to load a CSV file using Aspose.Cells for .NET while preserving numeric values as text. The sample creates a temporary CSV, disables numeric conversion via TxtLoadOptions.ConvertNumericData, reads cells with StringValue, and optionally saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | TxtLoadOptions | ConvertNumericData | load CSV as text | prevent numeric conversion | read numeric cells as string | preserve leading zeros | spreadsheet library
// Common Searches: Aspose.Cells load CSV without converting numbers | TxtLoadOptions ConvertNumericData false example | Read numeric values as text from CSV in C# | Preserve leading zeros when importing CSV with Aspose.Cells | How to keep numbers as strings in Aspose.Cells workbook
// Developer Intent: Load a CSV workbook and keep every numeric entry stored as a text string.
// Use Cases: Import product codes that contain leading zeros without losing formatting. | Maintain exact price strings for financial audits where rounding is unacceptable. | Validate raw data before any type conversion in data‑cleaning pipelines.
// AI Prompts: Generate C# code that loads a CSV with Aspose.Cells, disables numeric conversion, and prints each cell as a string. | Provide an Aspose.Cells example that uses TxtLoadOptions.ConvertNumericData = false to read a CSV and then saves it as an Excel file while preserving all values as text.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Shows how to load a CSV file using Aspose.Cells for .NET while preserving numeric values as text. The sample creates a temporary CSV, disables numeric conversion via TxtLoadOptions.ConvertNumericData, reads cells with StringValue, and optionally saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        // Sample CSV data containing numeric values
        string csvData = "ID,Price,Quantity\n1,19.99,5\n2,24.50,10";

        // Write the CSV data to a temporary file
        string tempCsvPath = Path.GetTempFileName();
        File.WriteAllText(tempCsvPath, csvData, Encoding.UTF8);

        // Create TxtLoadOptions and disable numeric conversion
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
        {
            ConvertNumericData = false   // Read numeric cells as text
        };

        // Load the CSV file using the specified options
        Workbook workbook = new Workbook(tempCsvPath, loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        // Access the loaded cells as strings (they remain text)
        Console.WriteLine("A2 (ID) as text: " + sheet.Cells["A2"].StringValue);
        Console.WriteLine("B2 (Price) as text: " + sheet.Cells["B2"].StringValue);
        Console.WriteLine("C2 (Quantity) as text: " + sheet.Cells["C2"].StringValue);

        // Save the workbook to an Excel file (optional)
        string outputPath = "LoadedAsText.xlsx";
        workbook.Save(outputPath);
    }
}
