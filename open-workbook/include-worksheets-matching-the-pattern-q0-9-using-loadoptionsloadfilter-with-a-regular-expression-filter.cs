using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

class RegexLoadFilter : LoadFilter
{
    private readonly Regex _regex;

    public RegexLoadFilter(string pattern)
    {
        _regex = new Regex(pattern, RegexOptions.Compiled);
    }

    public override void StartSheet(Worksheet sheet)
    {
        // Load full data only for sheets whose name matches the pattern "Q[0-9]+"
        if (_regex.IsMatch(sheet.Name))
        {
            LoadDataFilterOptions = LoadDataFilterOptions.All;
        }
        else
        {
            // Load only the sheet structure for non‑matching sheets (no cell data)
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}

class Program
{
    static void Main()
    {
        // Configure load options with the custom regex filter
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFilter = new RegexLoadFilter(@"^Q\d+$"),
            KeepUnparsedData = false // optional: improve performance when only reading
        };

        // Load the workbook using the filter
        Workbook workbook = new Workbook("template.xlsx", loadOptions);

        // Demonstrate which worksheets were loaded
        Console.WriteLine("Loaded worksheets:");
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Console.WriteLine(ws.Name);
        }

        // Save the filtered workbook if needed
        workbook.Save("filtered_output.xlsx");
    }
}