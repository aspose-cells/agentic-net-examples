// Title: Replace Invalid XML Characters When Loading CSV with Aspose.Cells for .NET – No Exceptions
// Description: Loads a UTF‑8 CSV into an Aspose.Cells Workbook after automatically substituting XML‑illegal characters with spaces. Uses TxtLoadOptions (CSV format) and a memory stream to prevent runtime exceptions, then saves the workbook as XLSX.
// Keywords: Aspose.Cells CSV load | replace invalid XML characters C# | remove illegal characters Excel | TxtLoadOptions CSV Aspose | sanitize CSV before XLSX conversion | .NET workbook creation | avoid exceptions Aspose.Cells
// Common Searches: how to clean CSV before loading with Aspose.Cells | Aspose.Cells replace invalid characters during CSV import | C# remove XML illegal characters from CSV | load CSV to Excel without exceptions Aspose | TxtLoadOptions CheckExcelRestriction false example
// Developer Intent: Automatically filter out characters that Excel/XML rejects while loading a CSV so the operation completes without throwing exceptions.
// Use Cases: Pre‑process user‑uploaded CSV files that contain control or unsupported Unicode symbols before converting them to XLSX. | Integrate a safe CSV‑to‑Excel pipeline in a .NET web service where data integrity must be preserved and errors avoided. | Batch‑convert legacy CSV datasets with hidden invalid characters into clean Excel workbooks using Aspose.Cells.
// AI Prompts: Generate C# code that scans a CSV string, replaces any character not allowed by XML 1.0 with a space, and loads the result into an Aspose.Cells Workbook via TxtLoadOptions. | Show how to configure TxtLoadOptions (separator, text qualifier, CheckExcelRestriction) to import a cleaned CSV from a MemoryStream and save it as XLSX without errors. | Explain why setting CheckExcelRestriction to false is required when sanitizing CSV data for Aspose.Cells and how it affects Excel‑specific validation.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads a UTF‑8 CSV into an Aspose.Cells Workbook after automatically substituting XML‑illegal characters with spaces. Uses TxtLoadOptions (CSV format) and a memory stream to prevent runtime exceptions, then saves the workbook as XLSX.
class CsvLoadReplaceInvalidChars
{
    static void Main()
    {
        // Paths for input CSV and output Excel file
        string sourceCsv = "input.csv";
        string outputXlsx = "output.xlsx";

        // Read the raw CSV content (assume UTF‑8 encoding)
        string rawContent = File.ReadAllText(sourceCsv, Encoding.UTF8);

        // Replace characters that are invalid for Excel/XML with a space
        string cleanedContent = RemoveInvalidXmlChars(rawContent);

        // Load the cleaned CSV data into a workbook using TxtLoadOptions
        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(cleanedContent)))
        {
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = ',',               // CSV delimiter
                HasTextQualifier = true,       // Enable text qualifier handling
                TextQualifier = '"',           // Default text qualifier
                CheckExcelRestriction = false // Allow otherwise restricted data
            };

            // No exception should be thrown here
            Workbook workbook = new Workbook(ms, loadOptions);

            // Save the workbook as an XLSX file
            workbook.Save(outputXlsx, SaveFormat.Xlsx);
        }

        Console.WriteLine("CSV loaded and saved without exceptions.");
    }

    // Removes characters that are not allowed in XML/Excel cell values
    static string RemoveInvalidXmlChars(string text)
    {
        StringBuilder sb = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            sb.Append(IsInvalidXmlChar(c) ? ' ' : c);
        }
        return sb.ToString();
    }

    // Determines whether a character is invalid according to XML 1.0 spec
    static bool IsInvalidXmlChar(char ch)
    {
        // Valid: #x9 | #xA | #xD | #x20‑#xD7FF | #xE000‑#xFFFD
        if (ch == 0x9 || ch == 0xA || ch == 0xD) return false;
        if (ch >= 0x20 && ch <= 0xD7FF) return false;
        if (ch >= 0xE000 && ch <= 0xFFFD) return false;
        return true;
    }
}
