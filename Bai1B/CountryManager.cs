using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Bai1B
{
    class CountryManager
    {
        List<Country> lstCountry { get; set; }

        public CountryManager() {
            lstCountry = new List<Country>();

        }
        public void loadFromFilePath(string pathFile) {

            if (File.Exists(pathFile))
            {
                var lines = File.ReadAllLines(pathFile);
                Country country = null;
                var lstLine = lines.Skip(8);
                foreach (var line in lstLine) {
                    var values = line.Split('=');
                    var value = values[1];
                    switch (values[0])
                    {
                        case "country":
                            if (country == null)
                            {
                                country = new Country();
                                country.country = int.Parse(value);
                            }
                            else
                            {
                                lstCountry.Add(country);
                                country = new Country();
                                country.country = int.Parse(value);

                            }
                            break;
                        case "name":
                            country.name = value;
                            break;
                        case "longName":
                            country.longName = value;
                            break;
                        case "foundingDate":
                            if (value.Contains('/') == true)
                            {
                                country.foundingDate = DateTime.ParseExact(value, "M/d/yyyy", new CultureInfo("en-US"));
                            }
                            else if (value.Contains('-') == true && value[0] != '-')
                            {
                                country.foundingDate = DateTime.ParseExact(value, "yyyy-MM-dd", new CultureInfo("en-US"));
                            }
                            else
                            {
                                
                            }
                           
                            
                            break;
                        case "population":
                            country.population = Int64.Parse(value);
                            break;
                        case "capital":
                            country.capital = value;
                            break;
                        case "largestCity":
                            country.largestCity = value;
                            break;
                            
                        case "area":
                            if (value.Contains("mi") == true)
                            {
                                var area = value.Split(new string[] { "mi" }, StringSplitOptions.None)[0];
                                if (area[0] == 'o') {
                                    area = area.Split(' ')[1];
                                }
                                country.area = Double.Parse(area) * 1.609344;
                            }
                            else {
                                var area = value.Split(new string[] { "km" }, StringSplitOptions.None)[0];
                                country.area = Double.Parse(area);
                            }
                            break;

                    }
                }
            }
        }
        public void exportToFilePath(string pathFile)
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                p.Workbook.Properties.Author = "tichondrius";
                p.Workbook.Properties.Title = "SkyNet Monthly Report";
                p.Workbook.Properties.Company = "hcmus is the best university";
                p.Workbook.Worksheets.Add("April 2012");
                ExcelWorksheet ws = p.Workbook.Worksheets[1]; 
                ws.Name = "April 2012";
                ws.Cells[1, 1].Value = "country";
                ws.Cells[1, 2].Value = "name";
                ws.Cells[1, 3].Value = "longName";
                ws.Cells[1, 4].Value = "foundingDate";
                ws.Cells[1, 5].Value = "population";
                ws.Cells[1, 6].Value = "capital";
                ws.Cells[1, 7].Value = "largestCity";
                ws.Cells[1, 8].Value = "area";

                ws.Cells[2, 1].LoadFromCollection<Country>(lstCountry);

                Byte[] bin = p.GetAsByteArray();
                File.WriteAllBytes(pathFile, bin);
            }
        }

        public void removeEmptyRecords()
        {

            lstCountry.RemoveAll(country => country.capital == null
                && country.area == null
                && country.longName == null
                && country.foundingDate == null
                && country.population == null
                && country.largestCity == null
                && country.name == null);
        }
        public void removeDuplicateInfo() {

            lstCountry = lstCountry
                .GroupBy(h => new { h.area, h.longName, h.foundingDate, h.population, h.largestCity, h.name })
                .Select(gr => gr.First())
                .ToList();
        }
    }
}
