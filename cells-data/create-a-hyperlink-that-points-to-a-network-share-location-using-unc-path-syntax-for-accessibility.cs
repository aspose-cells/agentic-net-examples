// Title: Add a UNC Network Share Hyperlink to an Excel Workbook with Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a hyperlink in cell A1 that points to a UNC path (\\Server\Share\Folder\File.txt), set custom display text, and save the file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# UNC hyperlink | Excel network share link | Hyperlinks.Add | UNC path | Aspose.Cells example | Excel hyperlink .NET | network share Excel
// Common Searches: Aspose.Cells add hyperlink to UNC path | C# create Excel link to network share | How to use Hyperlinks.Add for file server paths | Excel workbook with clickable network file link | Aspose.Cells UNC network share example
// Developer Intent: Insert a clickable link in an Excel worksheet that opens a file located on a network share using Aspose.Cells for .NET.
// Use Cases: Provide one‑click access to shared documents from a generated report. | Build a dashboard where cells open server‑based data files when selected. | Create a template that links to a central policy document stored on a shared drive.
// AI Prompts: Generate C# code with Aspose.Cells to add a UNC hyperlink and customize its display text. | Explain the Hyperlinks.Add method for linking to network shares and discuss any security considerations. | Show how to add multiple UNC hyperlinks to different cells in the same workbook using Aspose.Cells.

using Aspose.Cells;

// Shows how to create a workbook, insert a hyperlink in cell A1 that points to a UNC path (\\Server\Share\Folder\File.txt), set custom display text, and save the file using Aspose.Cells in C#.
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
