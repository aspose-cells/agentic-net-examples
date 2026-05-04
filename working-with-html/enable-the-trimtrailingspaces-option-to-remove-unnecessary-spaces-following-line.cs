using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class TrimTrailingSpacesCsv
{
    static void Main()
    {
        // Create a temporary CSV file that contains trailing spaces after line breaks
        string csvContent = "Name,Comment\nJohn,\"Hello   \"\nJane,\"World   \"\n";
        string tempCsvPath = "temp.csv";
        File.WriteAllText(tempCsvPath, csvContent, Encoding.UTF8);

        // Configure load options for CSV
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        loadOptions.Separator = ',';               // Use comma as delimiter
        loadOptions.HasTextQualifier = true;       // Enable text qualifier handling
        loadOptions.TextQualifier = '\"';          // Set double‑quote as the qualifier

        // Load the CSV file into a workbook
        Workbook workbook = new Workbook(tempCsvPath, loadOptions);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Iterate through all cells and trim trailing spaces from string values
        foreach (Cell cell in cells)
        {
            if (cell.Type == CellValueType.IsString)
            {
                string original = cell.StringValue;
                string trimmed = original.TrimEnd(); // Remove spaces at the end
                if (trimmed != original)
                {
                    cell.PutValue(trimmed);
                }
            }
        }

        // Save the processed workbook back to CSV (in memory) to verify the result
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Separator = ','
        };

        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, saveOptions);
            string resultCsv = Encoding.UTF8.GetString(ms.ToArray());
            Console.WriteLine("Trimmed CSV output:");
            Console.WriteLine(resultCsv);
        }

        // Clean up temporary file
        File.Delete(tempCsvPath);
    }
}