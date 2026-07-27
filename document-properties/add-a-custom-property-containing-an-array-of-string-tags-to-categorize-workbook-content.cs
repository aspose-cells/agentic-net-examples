using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsTagPropertyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Define an array of tags
            string[] tags = new[] { "Finance", "Quarterly", "Report" };

            // Convert the array to a JSON‑like string representation
            // (Custom document properties accept only scalar values, so we store the array as a string)
            string tagsValue = "[" + string.Join(",", Array.ConvertAll(tags, t => $"\"{t}\"")) + "]";

            // Add the custom document property named "Tags" with the stringified array
            // (uses the Add(string, string) overload from CustomDocumentPropertyCollection)
            workbook.CustomDocumentProperties.Add("Tags", tagsValue);

            // Save the workbook (lifecycle rule)
            workbook.Save("WorkbookWithTags.xlsx");
        }
    }
}