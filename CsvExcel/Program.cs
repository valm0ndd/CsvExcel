using CsvExcel.Model;
using OfficeOpenXml;
using System.Globalization;

namespace CSVToExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string inputFile = Path.Combine(desktopPath, "CSV", "people.csv");
            string outputFile = Path.Combine(desktopPath, "CSV", "filtered-people.xlsx");
            List<PersonModel> outputRecords = new List<PersonModel>();

            string filterLetter = string.Empty;
            while (true)
            {
                Console.Write("Enter the first letter of the name: ");
                filterLetter = Console.ReadLine();
                if (!string.IsNullOrWhiteSpace(filterLetter) && filterLetter.Length == 1 && char.IsLetter(filterLetter[0]))
                {
                    break;
                }
                Console.WriteLine("The input must not be empty, must not contain more than 1 letter and must not contain numbers!");
                Console.WriteLine("Please enter a single letter: ");
            }

            using (var reader = new StreamReader(inputFile))
            using (var csv = new CsvHelper.CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<PersonModel>();
                foreach (var record in records)
                {
                    if (record.FirstName.StartsWith(filterLetter, StringComparison.OrdinalIgnoreCase))
                    {
                        outputRecords.Add(record);
                        Console.WriteLine(record.FirstName + " " + record.LastName);
                    }
                }
            }

            Console.WriteLine($"Number of records starting with {filterLetter.ToUpper()}: {outputRecords.Count}");

            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Filtered People");
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Filtered People");
                }
                else
                {
                    if (worksheet.Dimension != null)
                    {
                        int startRow = 2;
                        int endRow = worksheet.Dimension.End.Row;
                        if (endRow >= startRow)
                        {
                            string clearRange = $"A{startRow}:Z{endRow}";
                            worksheet.Cells[clearRange].Clear();
                        }
                    }
                }

                worksheet.Cells["A1"].Value = "Index";
                worksheet.Cells["B1"].Value = "User Id";
                worksheet.Cells["C1"].Value = "First Name";
                worksheet.Cells["D1"].Value = "Last Name";
                worksheet.Cells["E1"].Value = "Sex";
                worksheet.Cells["F1"].Value = "Email";
                worksheet.Cells["G1"].Value = "Phone";
                worksheet.Cells["H1"].Value = "Date of Birth";
                worksheet.Cells["I1"].Value = "Job Title";

                if (outputRecords.Any())
                {
                    worksheet.Cells["A2"].LoadFromCollection(outputRecords, false);
                }
                else if (worksheet.Dimension != null)
                {
                    worksheet.Cells["A2:Z" + worksheet.Dimension.End.Row].Clear();
                }

                package.Save();
            }

            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}