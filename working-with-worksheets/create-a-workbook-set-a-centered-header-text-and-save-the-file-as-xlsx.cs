// Title: Aspose.Cells .NET: Create an XLSX workbook with a centered header
// Description: Demonstrates how to generate a new XLSX workbook using Aspose.Cells for .NET, access the first worksheet, configure the PageSetup to add a centered header, and save the file as "CenteredHeader.xlsx".
// Keywords: Aspose.Cells .NET | create XLSX workbook | centered header | PageSetup SetHeader | Excel header formatting | save workbook as XLSX | C# Aspose.Cells example
// Common Searches: Aspose.Cells set centered header C# | how to add a header to Excel with Aspose.Cells | C# create XLSX file and set page header | Aspose.Cells PageSetup header example | center header Aspose.Cells .NET
// Developer Intent: Add a centered header to a newly created XLSX workbook and persist the file.
// Use Cases: Produce printable reports where each page displays a centered title in the header. | Build a corporate template that automatically inserts a centered document name on every printed sheet. | Automate bulk workbook generation with a uniform centered header for branding consistency.
// AI Prompts: Generate C# code with Aspose.Cells that creates an XLSX file and inserts a centered header containing the current date. | Show how to set left, center, and right headers differently in Aspose.Cells and save the workbook as XLSX. | Explain how to customize font style, size, and color of a centered header using Aspose.Cells PageSetup.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderExample
{
    // Demonstrates how to generate a new XLSX workbook using Aspose.Cells for .NET, access the first worksheet, configure the PageSetup to add a centered header, and save the file as "CenteredHeader.xlsx".
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the centered header (section 1 = center) with desired text
            pageSetup.SetHeader(1, "My Centered Header Text");

            // Save the workbook as an XLSX file
            workbook.Save("CenteredHeader.xlsx");

            Console.WriteLine("Workbook created with centered header and saved as XLSX.");
        }
    }
}
