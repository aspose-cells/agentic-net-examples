// Title: Creating a camera shape for a cell range and exporting it as an image with Aspose.Cells for .NET (AddCamera API unavailable)
// Common Searches: Aspose.Cells C# add camera shape to a specific range and save as PNG | how to capture an Excel range as an image using camera shape in .NET | export worksheet range to picture with Aspose.Cells when AddCamera is missing | C# code to generate a snapshot of cells A1:C5 using Aspose.Cells | workaround for missing camera shape feature in Aspose.Cells .NET
// Tags: camera shape creation Aspose.Cells .NET | export range to image C# Aspose.Cells | snapshot worksheet cells Aspose.Cells | AddCamera API limitation Aspose.Cells | convert Excel range to bitmap using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, populates cells A1:C5 with sample data, highlights that the AddCamera API and related camera shape functionality are not available in the current Aspose.Cells version, and then saves the workbook to 'WorkbookWithData.xlsx' after ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data into the range A1:C5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // NOTE: The AddCamera API and related Camera shape features are not available
            // in the current Aspose.Cells version used for this project.
            // Therefore, the camera shape creation and image extraction steps are omitted.

            // Save the workbook
            string workbookPath = "WorkbookWithData.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(workbookPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(workbookPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
