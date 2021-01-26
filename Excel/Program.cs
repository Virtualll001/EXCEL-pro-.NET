using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using LicenseContext = OfficeOpenXml.LicenseContext;

//zdroj: https://www.youtube.com/watch?v=j3S3aI8nMeE

namespace Excel
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Demos\YouTubeDemo.xlsx");
            var people = GetSetupData();
            await SaveExcelFile(people, file);

            //Čtení z Excelu
            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            foreach (var p in peopleFromExcel)
            {
                Console.WriteLine($"{p.Id} {p.Krestni} {p.Prijmeni}");
            }
            Console.ReadKey();
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var ws = package.Workbook.Worksheets[0];
            int row = 3; //chceme data = přeskočíme nadpis tabulky na 1. řádku a nadpis sloupce na 2. řádku
            int col = 1;

            //hledá dokud v prvním sloupci (required [Key] s Id nebude prázdno)
            while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
            {
                PersonModel p = new();
                p.Id = int.Parse(ws.Cells[row, col].Value.ToString()); //přes string převedeme double na int
                p.Krestni = ws.Cells[row,col + 1].Value.ToString();
                p.Prijmeni = ws.Cells[row, col + 2].Value.ToString();
                output.Add(p);
                row += 1;
            }
            return output;
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);
            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Report"); //nový list v xls
            var range = ws.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            //formátování
            ws.Cells["A1"].Value = "Seznam jmen";
            ws.Cells["A1:C1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;
            ws.Column(3).Width = 20;

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }

        private static List<PersonModel> GetSetupData() //testovací data
        {
            List<PersonModel> output = new()
            {
                new() { Id = 1, Krestni = "Luboš", Prijmeni = "Harwey" },
                new() { Id = 2, Krestni = "Tomáš", Prijmeni = "Jihlavský" },
                new() { Id = 3, Krestni = "Jane", Prijmeni = "Novotná" }
            };
            return output;
        }
    }
    public class PersonModel
    {
        [Key]
        public int Id { get; set; }
        [DisplayName("Křestní jméno")]
        public string Krestni { get; set; }
        [DisplayName("Příjmení")]
        public string Prijmeni { get; set; }
    }
}
