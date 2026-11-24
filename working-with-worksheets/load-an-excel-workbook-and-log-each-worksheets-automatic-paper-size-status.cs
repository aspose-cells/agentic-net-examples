// Title: Load an Excel workbook and log each worksheet's automatic paper size flag using Aspose.Cells for .NET
// Common Searches: how to check if a worksheet uses automatic page size with Aspose.Cells in C# | Aspose.Cells get IsAutomaticPageSize property for all sheets | C# read worksheet page setup automatic paper size using Aspose.Cells | load workbook and enumerate page setup settings with Aspose.Cells .NET | handle missing Excel file when opening workbook using Aspose.Cells
// Tags: Aspose.Cells worksheet page setup automatic size | C# load Excel workbook Aspose.Cells | enumerate worksheets page setup properties | reflection fallback for IsAutomaticPageSize | create empty workbook if file not found Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads "input.xlsx" with Aspose.Cells, creates a new empty workbook when the file is absent, then iterates through all worksheets. For each sheet it uses reflection to read the PageSetup.IsAutomaticPageSize boolean and prints the worksheet name together with its automatic paper size status.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        Workbook workbook = null;

        try
        {
            // Ensure the input file exists; if not, create an empty workbook to avoid FileNotFoundException
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"File '{inputPath}' not found. Creating a new empty workbook.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Iterate through each worksheet and report automatic paper size status
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            bool isAutomaticPaperSize = GetAutomaticPageSize(sheet);
            Console.WriteLine($"Worksheet '{sheet.Name}': Automatic Paper Size = {isAutomaticPaperSize}");
        }
    }

    private static bool GetAutomaticPageSize(Worksheet sheet)
    {
        try
        {
            var pageSetup = sheet.PageSetup;
            var propInfo = pageSetup.GetType().GetProperty("IsAutomaticPageSize");
            if (propInfo != null && propInfo.PropertyType == typeof(bool))
            {
                return (bool)propInfo.GetValue(pageSetup);
            }
        }
        catch
        {
            // Ignored – fallback to false if any reflection error occurs
        }

        // Default value when the property is unavailable
        return false;
    }
}
