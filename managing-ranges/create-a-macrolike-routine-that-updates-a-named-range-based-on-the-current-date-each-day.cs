// Title: Aspose.Cells for .NET – Macro‑Like Routine to Refresh a Named Range with Today’s Date
// Description: A concise C# example that creates (or opens) a workbook, defines a named range, writes the current system date into every cell of that range, and saves the file. The routine can be scheduled to run daily, providing a macro‑free way to keep a named range up‑to‑date.
// Keywords: Aspose.Cells update named range | C# set today date in Excel | macro‑like routine Aspose.Cells | refresh named range daily .NET | populate date in named range C# | Aspose.Cells US developers | Excel automation Europe | global Excel date refresh
// Common Searches: how to set a named range to today's date using Aspose.Cells | C# code to refresh an Excel named range each day | Aspose.Cells macro alternative for date update | schedule Aspose.Cells routine to write current date | update existing named range with current date .NET
// Developer Intent: Provide a reusable .NET method that writes the current date into a specified named range so the workbook always reflects today’s date when the routine runs.
// Use Cases: Daily report templates that display the report date in a header cell via a named range. | Automated data feeds where a shared workbook must show the processing date without manual edits. | Windows Task Scheduler or Azure Function that executes the routine each morning to keep a master workbook current.
// AI Prompts: Generate C# Aspose.Cells code that opens an existing workbook, finds a named range, and updates all its cells to DateTime.Today. | Explain how to integrate the UpdateNamedRangeWithCurrentDate method into a Windows Service that runs at 06:00 UTC. | Show how to modify the sample so the named range can span multiple cells and still receive the current date.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroLike
{
    // A concise C# example that creates (or opens) a workbook, defines a named range, writes the current system date into every cell of that range, and saves the file. The routine can be scheduled to run daily, providing a macro‑free way to keep a named range up‑to‑date.
    public class UpdateNamedRangeWithCurrentDate
    {
        // Entry point for the console application
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Define a cell that will hold the current date (initial value)
            sheet.Cells["A1"].PutValue(DateTime.Today);

            // Create a named range that refers to the cell A1
            int nameIndex = workbook.Worksheets.Names.Add("TodayDate");
            Name todayName = workbook.Worksheets.Names[nameIndex];
            todayName.RefersTo = "=Sheet1!$A$1";

            // Retrieve the range associated with the named range
            Aspose.Cells.Range dateRange = todayName.GetRange();

            // Update every cell in the range to the current date
            foreach (Cell cell in dateRange)
            {
                cell.PutValue(DateTime.Today);
            }

            // Save the workbook (macro‑like routine can be called daily to refresh the date)
            string outputPath = "UpdatedNamedRange.xlsx";
            workbook.Save(outputPath);
        }
    }
}
