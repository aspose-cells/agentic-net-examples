using System;
using System.Text;
using Aspose.Cells;

namespace PrecedentsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Set a formula that references several cells, a range, and an external workbook
            cells["A1"].Formula = "=B1+SUM(B2:B5)+[ExternalBook.xlsx]Sheet1!C3";

            // Retrieve all precedent references of the formula cell A1
            ReferredAreaCollection precedents = cells["A1"].GetPrecedents();

            // If there are precedents, iterate and log their addresses
            if (precedents != null && precedents.Count > 0)
            {
                Console.WriteLine("Precedent references for cell A1:");
                foreach (ReferredArea area in precedents)
                {
                    StringBuilder sb = new StringBuilder();

                    // Include external file name if the reference is an external link
                    if (area.IsExternalLink)
                    {
                        sb.Append('[').Append(area.ExternalFileName).Append(']');
                    }

                    // Append sheet name
                    sb.Append(area.SheetName).Append('!');

                    // Append start cell address
                    sb.Append(CellsHelper.CellIndexToName(area.StartRow, area.StartColumn));

                    // If the reference is a range, append the end cell address
                    if (area.IsArea)
                    {
                        sb.Append(':')
                          .Append(CellsHelper.CellIndexToName(area.EndRow, area.EndColumn));
                    }

                    Console.WriteLine(sb.ToString());
                }
            }
            else
            {
                Console.WriteLine("No precedents found for cell A1.");
            }

            // Optionally save the workbook (demonstrates lifecycle rule usage)
            workbook.Save("PrecedentsDemo.xlsx");
        }
    }
}