// Title: Handle unsupported CultureInfo identifiers in Aspose.Cells LoadOptions with fallback logic
// Description: Demonstrates how to protect workbook loading from invalid locale strings by catching CultureNotFoundException, using a custom SafeCultureInfoFactory, and falling back to InvariantCulture or a default culture before creating the Workbook.
// Keywords: Aspose.Cells LoadOptions CultureInfo | CultureNotFoundException handling | fallback culture Aspose.Cells | invalid LCID .NET | custom implementation factory | globalization Excel loading | C# workbook locale error handling | InvariantCulture fallback
// Common Searches: Aspose.Cells load workbook with invalid culture | catch CultureNotFoundException in LoadOptions | provide default CultureInfo for unsupported locale | custom factory for CultureInfo in Aspose.Cells | how to use SafeCultureInfoFactory Aspose.Cells | set invariant culture when loading Excel file
// Developer Intent: Prevent runtime failures when an unsupported locale identifier is supplied to LoadOptions.CultureInfo by implementing graceful fallback mechanisms.
// Use Cases: User‑entered applications where the locale is entered dynamically and may be invalid. | Enterprise services that process Excel files from multiple regions and need a reliable default culture. | Automated pipelines that must continue processing even when a specific LCID is not installed on the host machine.
// AI Prompts: Generate C# code that assigns LoadOptions.CultureInfo from a string and defaults to CultureInfo.InvariantCulture on error using Aspose.Cells. | Create a custom implementation factory for Aspose.Cells that returns a fallback CultureInfo when an LCID is not supported. | Show how to log and handle CultureNotFoundException during workbook loading while preserving existing LoadOptions settings.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureInfoErrorHandling
{
    // Optional: custom factory to provide a fallback CultureInfo when the LCID is unsupported
    // Demonstrates how to protect workbook loading from invalid locale strings by catching CultureNotFoundException, using a custom SafeCultureInfoFactory, and falling back to InvariantCulture or a default culture before creating the Workbook.
    public class SafeCultureInfoFactory : CustomImplementationFactory
    {
        public override CultureInfo CreateCultureInfo(int lcid)
        {
            try
            {
                // Attempt to create the requested CultureInfo
                return base.CreateCultureInfo(lcid);
            }
            catch (CultureNotFoundException)
            {
                // Fallback to invariant culture if the LCID is not supported
                Console.WriteLine($"LCID {lcid} is not supported. Using InvariantCulture instead.");
                return CultureInfo.InvariantCulture;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Register the custom factory (optional but demonstrates a global fallback)
            CellsHelper.CustomImplementationFactory = new SafeCultureInfoFactory();

            // Prepare LoadOptions for loading an XLSX file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Attempt to assign an unsupported CultureInfo identifier
            try
            {
                // This will throw CultureNotFoundException for an invalid culture name
                loadOptions.CultureInfo = new CultureInfo("xx-XX");
            }
            catch (CultureNotFoundException ex)
            {
                // Handle the error and fall back to a known culture (e.g., invariant or en-US)
                Console.WriteLine($"Unsupported culture identifier: {ex.InvalidCultureName}");
                loadOptions.CultureInfo = CultureInfo.InvariantCulture;
            }

            // Path to the source workbook (replace with an actual file path)
            string sourcePath = "sample.xlsx";

            // Load the workbook using the prepared LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Example operation: read a cell value to verify loading succeeded
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Cell A1 value after loading with culture '{loadOptions.CultureInfo.Name}': {cellValue}");

            // Save the workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
