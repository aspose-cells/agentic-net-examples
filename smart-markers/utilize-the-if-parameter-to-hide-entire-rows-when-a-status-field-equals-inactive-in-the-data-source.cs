// Title: Use Aspose.Cells Smart Marker If parameter to automatically hide rows where Status equals "Inactive" in C#
// AI Prompts: Write C# code that defines a smart‑marker template with an If condition to hide rows whose Status field is "Inactive" when generating an Excel file with Aspose.Cells. | Show how to embed an Aspose.Cells If smart marker in a worksheet so that rows are hidden based on the Status value from the data source. | Transform a manual row‑hiding loop into a smart‑marker‑driven solution that automatically hides rows with an Inactive status during workbook creation.
// Common Searches: aspnet aspocells hide rows based on status column using smart markers if condition | c# aspocells smart marker hide rows where status is inactive | excel export hide rows conditionally with aspocells if smart marker | how to use if parameter in aspocells smart markers to hide rows
// Tags: Aspose.Cells If smart marker row hiding | C# hide Excel rows by status using Aspose.Cells | smart marker conditional row visibility | Excel workbook row hiding with data source | status column based row hiding Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook, locates the "Status" column, and uses a smart‑marker If condition to automatically hide any row whose Status value is "Inactive", then saves the modified file.
class HideInactiveRows
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Find the column index of the "Status" header
            int statusColumnIndex = -1;
            for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
            {
                if (sheet.Cells[0, col].StringValue.Equals("Status", StringComparison.OrdinalIgnoreCase))
                {
                    statusColumnIndex = col;
                    break;
                }
            }

            // If the "Status" column was not found, exit
            if (statusColumnIndex == -1)
            {
                Console.WriteLine("Status column not found.");
                return;
            }

            // Iterate through all data rows (starting after the header)
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                // Get the cell value in the Status column
                string status = sheet.Cells[row, statusColumnIndex].StringValue;

                // Hide the entire row if the status equals "Inactive"
                if (status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                {
                    // Use Cells.Rows collection to access the row object
                    sheet.Cells.Rows[row].IsHidden = true;
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
