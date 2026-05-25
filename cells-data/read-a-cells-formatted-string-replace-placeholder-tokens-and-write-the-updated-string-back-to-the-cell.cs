using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a formatted string containing placeholder tokens into cell A1
        worksheet.Cells["A1"].PutValue("Dear {Name}, your balance is {Balance} USD.");

        // Configure replace options (case‑insensitive, partial match)
        ReplaceOptions replaceOptions = new ReplaceOptions
        {
            CaseSensitive = false,
            MatchEntireCellContents = false
        };

        // Replace the {Name} placeholder
        worksheet.Cells["A1"].Replace("{Name}", "John Doe", replaceOptions);

        // Replace the {Balance} placeholder
        worksheet.Cells["A1"].Replace("{Balance}", "1234.56", replaceOptions);

        // Output the final formatted string to the console (optional)
        Console.WriteLine(worksheet.Cells["A1"].StringValue);

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}