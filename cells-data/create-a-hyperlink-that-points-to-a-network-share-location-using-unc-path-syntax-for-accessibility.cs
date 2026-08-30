// Title: Insert a UNC network share hyperlink into an Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a hyperlink to cell A1 pointing to a UNC path like \\Server\Share\Folder\File.txt and sets a custom display text. | Show how to create and save an Excel workbook that contains a network‑share hyperlink using Aspose.Cells, including configuring the hyperlink address and visible caption.
// Common Searches: how to add a UNC path hyperlink to an Excel cell using Aspose.Cells C# | Aspose.Cells C# create hyperlink to network share and set display text | C# Aspose.Cells save workbook with hyperlink to \\Server\Share | example of adding a file share hyperlink in Excel with Aspose.Cells .NET | Aspose.Cells hyperlink to network location in .xlsx file
// Tags: Aspose.Cells create UNC link | C# Excel hyperlink to network share | Aspose.Cells set hyperlink caption | save workbook with external file link | hyperlink address specification Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, adds a hyperlink in cell A1 that points to the UNC network share "\\Server\Share\Folder\File.txt", sets the cell's display text to "Open Network File", and saves the workbook as "NetworkHyperlink.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // UNC path to the network share
        string uncPath = @"\\Server\Share\Folder\File.txt";

        // Add a hyperlink to cell A1 that points to the UNC location
        worksheet.Hyperlinks.Add("A1", 1, 1, uncPath);

        // Set the display text for the hyperlink
        worksheet.Cells["A1"].PutValue("Open Network File");

        // Save the workbook
        workbook.Save("NetworkHyperlink.xlsx");
    }
}
