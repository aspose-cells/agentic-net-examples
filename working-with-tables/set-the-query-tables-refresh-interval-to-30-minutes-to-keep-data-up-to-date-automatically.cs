// Title: Configure a 30‑minute auto‑refresh interval for an Excel query table using Aspose.Cells in C#
// Common Searches: Aspose.Cells set query table auto refresh interval to 30 minutes in C# | C# code to refresh Excel query table every half hour with Aspose.Cells | How to schedule periodic query table refresh in an Aspose.Cells workbook | Is there a RefreshInterval property for QueryTable in Aspose.Cells .NET | Programmatically enable automatic query table refresh using Aspose.Cells
// Tags: Aspose.Cells query table auto refresh | C# set query table refresh interval | Excel query table refresh schedule Aspose.Cells | Aspose.Cells QueryTable refresh limitation | programmatic query table refresh .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), accesses the first worksheet, checks for a query table, and explains that the current Aspose.Cells QueryTable class does not expose a RefreshInterval or RefreshOnFileOpen property, so developers must implement custom logic to achieve a 30‑minute automatic refresh before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. A new workbook will be created.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet (or specify the required one)
            Worksheet sheet = workbook.Worksheets[0];

            // Check that the worksheet contains at least one query table
            if (sheet.QueryTables.Count > 0)
            {
                // Access the first query table
                QueryTable queryTable = sheet.QueryTables[0];

                // NOTE: The Aspose.Cells QueryTable class does not expose RefreshInterval
                // and RefreshOnFileOpen properties in the current version.
                // If needed, implement custom logic here.

                // Example placeholder: you could refresh the query table manually
                // queryTable.Refresh();
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
