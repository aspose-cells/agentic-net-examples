using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class CsvLoadReplaceInvalidChars
{
    static void Main()
    {
        // Sample CSV content containing control characters that are invalid in Excel cells
        string rawCsv = "Name,Age,Note\nJohn,30,\"Good\u0000Student\"\nJane,25,\"Excellent\u0001Student\"";

        // Replace invalid characters (control chars except CR/LF) with a space
        string cleanedCsv = CleanInvalidCharacters(rawCsv);

        // Load the cleaned CSV data from a memory stream
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(cleanedCsv)))
        {
            // Configure CSV load options
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = ',',               // Column delimiter
                HasTextQualifier = true,       // Enable text qualifier handling
                TextQualifier = '\"',          // Default text qualifier
                ConvertNumericData = true,     // Convert numeric strings to numbers
                ConvertDateTimeData = true,    // Convert date strings to DateTime
                CheckExcelRestriction = false // Prevent Excel restriction checks during load
            };

            Workbook workbook = null;
            try
            {
                // Load workbook using the cleaned CSV stream and options
                workbook = new Workbook(csvStream, loadOptions);
                Console.WriteLine("CSV loaded successfully without exceptions.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during CSV load: {ex.Message}");
                return;
            }

            // Access the first worksheet and display some cell values
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"A1: {sheet.Cells["A1"].StringValue}");
            Console.WriteLine($"C2: {sheet.Cells["C2"].StringValue}");

            // Save the workbook to an Excel file
            workbook.Save("CleanedOutput.xlsx", SaveFormat.Xlsx);
        }
    }

    // Helper method that replaces control characters (except CR/LF) with a space
    static string CleanInvalidCharacters(string input)
    {
        var sb = new StringBuilder(input.Length);
        foreach (char c in input)
        {
            if (c == '\r' || c == '\n' || !char.IsControl(c))
                sb.Append(c);
            else
                sb.Append(' ');
        }
        return sb.ToString();
    }
}