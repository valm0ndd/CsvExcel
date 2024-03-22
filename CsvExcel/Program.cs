using CsvExcel.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace CSVToExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string sourceFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CSV");
            string destinationFolderPath = Path.Combine(desktopPath, "CSV-Filtered");

            if (Directory.Exists(destinationFolderPath))
            {
                Directory.Delete(destinationFolderPath, true);
            }

            Directory.CreateDirectory(destinationFolderPath);
            
            string inputFile = Path.Combine(sourceFolderPath, "people.csv");
            string outputFile = Path.Combine(destinationFolderPath, "filtered-people.xlsx");

            List<PersonModel> outputRecords = new List<PersonModel>();

            string filterLetter = GetFilterLetterFromUser();

            ReadAndFilterRecords(inputFile, filterLetter, outputRecords);

            WriteRecordsToExcel(outputFile, outputRecords);

            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }

        static string GetFilterLetterFromUser()
        {
            string filterLetter = string.Empty;
            while (true)
            {
                Console.Write("Enter the first letter of the name: ");
                filterLetter = Console.ReadLine()?.Trim().ToUpper();

                if (!string.IsNullOrWhiteSpace(filterLetter) && filterLetter.Length == 1 && char.IsLetter(filterLetter[0]))
                {
                    break;
                }

                Console.WriteLine("Invalid input! Please enter a single letter.");
            }
            return filterLetter;
        }
        static void ReadAndFilterRecords(string inputFile, string filterLetter, List<PersonModel> outputRecords)
        {
            try
            {
                using (var reader = new StreamReader(inputFile))
                using (var csv = new CsvHelper.CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<PersonModel>();
                    foreach (var record in records)
                    {
                        if (record.FirstName.StartsWith(filterLetter, StringComparison.OrdinalIgnoreCase))
                        {
                            outputRecords.Add(record);
                            Console.WriteLine($"{record.FirstName} {record.LastName}");
                        }
                    }
                }
                Console.WriteLine($"Number of records starting with {filterLetter.ToUpper()}: {outputRecords.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while reading the CSV file: {ex.Message}");
            }
        }
        static void WriteRecordsToExcel(string outputFile, List<PersonModel> outputRecords)
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while writing to Excel file: {ex.Message}");
            }
        }
    }
}
