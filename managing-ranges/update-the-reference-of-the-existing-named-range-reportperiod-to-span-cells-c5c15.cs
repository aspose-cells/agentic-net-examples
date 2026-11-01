// Title: How to update an existing named range to the cell range C5:C15 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change named range address to C5:C15 | Set RefersTo property of a named range with Aspose.Cells .NET | Programmatically update an existing named range in an Excel workbook using Aspose.Cells | Modify named range reference in a workbook with Aspose.Cells for C# | Update named range cells in an Aspose.Cells workbook
// Tags: Aspose.Cells modify named range reference | Aspose.Cells RefersTo property C# | C# update Excel named range Aspose.Cells | Aspose.Cells workbook named range management | change Excel named range address .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, retrieves the named range "ReportPeriod", updates its RefersTo property to the address $C$5:$C$15 on the first worksheet, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Retrieve the named range "ReportPeriod"
            Name reportPeriod = workbook.Worksheets.Names["ReportPeriod"];
            if (reportPeriod != null)
            {
                // Use the first worksheet (or any appropriate worksheet) for the reference
                Worksheet ws = workbook.Worksheets[0];

                // Update the reference to span cells C5:C15 on that worksheet
                reportPeriod.RefersTo = $"'{ws.Name}'!$C$5:$C$15";
            }
            else
            {
                Console.WriteLine("Named range \"ReportPeriod\" not found.");
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
