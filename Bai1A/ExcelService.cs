﻿using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bai1A
{
    class ExcelService
    {
        public static ExcelData getRecordFromPath(String inputPath)
        {
            ExcelData excelData = new ExcelData();
            using (var stream = File.Open(inputPath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    var columnsLength = reader.FieldCount;
                    Console.WriteLine(reader.ResultsCount);
                    int row = 0;
                    while (reader.Read())
                    {
                        var record = new ExpandoObject() as IDictionary<string, Object>;
                        for (int i = 0; i < columnsLength; i++)
                        {
                            if (row == 0)
                            {
                                var header = new TitleHeader();
                                header.Field = reader.GetString(i);
                                excelData.headers.Add(header);

                            }
                            else
                            {
                                if (reader.IsDBNull(i) == false)
                                {
                                    var type = reader.GetFieldType(i).Name;
                                    switch (type)
                                    {
                                        case "String":
                                            record.Add(excelData.headers[i].Field, reader.GetString(i));
                                            if (excelData.headers[i].TypeValue == null) {
                                                excelData.headers[i].TypeValue = "String";
                                            }
                                            break;
                                        case "Double": 
                                            record.Add(excelData.headers[i].Field, reader.GetDouble(i));
                                            if (excelData.headers[i].TypeValue == null) {
                                                excelData.headers[i].TypeValue = "Number";
                                            }
                                            break;
                                    }
                                }
                                else
                                {
                                    record.Add(excelData.headers[i].Field, null);
                                }
                            }
                        }
                        if (row != 0)
                        {
                            excelData.lstRecords.Add(record as ExpandoObject);
                        }
                        row++;
                    }

                }
            }
            return excelData;
        }
        public static void ExportExcelInPath(String Path, ExcelData excelData)
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                p.Workbook.Properties.Author = "Miles Dyson";
                p.Workbook.Properties.Title = "SkyNet Monthly Report";
                p.Workbook.Properties.Company = "Cyberdyne Systems";

                // The rest of our code will go here...
                p.Workbook.Worksheets.Add("April 2012");
                ExcelWorksheet ws = p.Workbook.Worksheets[1]; // 1 is the position of the worksheet
                ws.Name = "April 2012";

                var lenghtHeader = excelData.headers.Count();
                var lenghRecord = excelData.lstRecords.Count();

                //Write header
                for (int i = 0; i < lenghtHeader; i += 1)
                {
                    ws.Cells[1, i + 1].Value = excelData.headers[i].Field;
                }
                //Write record
                for (int i = 0; i < lenghRecord; i += 1)
                {
                    for (int j = 0; j < lenghtHeader; j += 1)
                    {
                        var value = excelData.lstRecords[i] as IDictionary<string, object>;
                        if (value != null)
                        {
                            ws.Cells[i + 2, j + 1].Value = value[excelData.headers[j].Field];
                        }
                    }
                }

                Byte[] bin = p.GetAsByteArray();
                File.WriteAllBytes(Path, bin);
            }
        }
    }
}
