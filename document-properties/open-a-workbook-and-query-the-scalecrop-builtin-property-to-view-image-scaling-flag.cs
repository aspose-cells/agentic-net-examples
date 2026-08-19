// Title: Read the ScaleCrop built‑in document property from an Excel workbook using Aspose.Cells for .NET
// Description: This C# example shows how to load an existing .xlsx file with Aspose.Cells, access the workbook's BuiltInDocumentPropertyCollection, retrieve the boolean ScaleCrop flag that indicates whether the thumbnail image is scaled and cropped, and output the value to the console.
// Keywords: Aspose.Cells ScaleCrop | C# read built‑in document property | Excel thumbnail scaling flag | Workbook built‑in properties .NET | Get ScaleCrop value Aspose | Aspose.Cells API example | C# Excel document properties
// Common Searches: how to get ScaleCrop property with Aspose.Cells | Aspose.Cells read built‑in document properties C# | retrieve thumbnail scaling flag from Excel file | ScaleCrop flag Aspose.Cells .NET example | access built‑in properties of a workbook using C#
// Developer Intent: Obtain the ScaleCrop flag from an Excel workbook to know if its thumbnail is scaled and cropped.
// Use Cases: Log the ScaleCrop setting for compliance before distributing a workbook. | Decide whether to regenerate a thumbnail after content changes based on the flag. | Trigger image‑processing logic only when the workbook thumbnail is set to be scaled and cropped.
// AI Prompts: Generate C# code with Aspose.Cells that sets the ScaleCrop property to true and saves the workbook. | Show how to enumerate all built‑in document properties of a workbook and print each name/value pair using Aspose.Cells. | Provide a snippet that checks the ScaleCrop flag and updates the workbook thumbnail accordingly before saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This C# example shows how to load an existing .xlsx file with Aspose.Cells, access the workbook's BuiltInDocumentPropertyCollection, retrieve the boolean ScaleCrop flag that indicates whether the thumbnail image is scaled and cropped, and output the value to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Query the ScaleCrop property (true if thumbnail is scaled and cropped)
        bool scaleCrop = properties.ScaleCrop;

        // Output the current value of ScaleCrop
        Console.WriteLine("ScaleCrop property value: " + scaleCrop);
    }
}
