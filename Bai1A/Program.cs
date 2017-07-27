using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bai1A
{
    class Program
    {

        static ExcelData getRecordFromPath(String inputPath)
        {
            ExcelData excelData = new ExcelData();
            using (var stream = File.Open("data.xlsx", FileMode.Open, FileAccess.Read))
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
                                    if (row >= 1)
                                    {
                                       
                                        var type = reader.GetFieldType(i).Name;
                                        switch (type) {
                                            case "String":
                                                excelData.headers[i].TypeValue = "String";
                                                if (reader.IsDBNull(i) == true) {
                                                    record.Add(excelData.headers[i].Field, null);
                                                }
                                                record.Add(excelData.headers[i].Field, reader.GetString(i));
                                                break;
                                            case "Integer":
                                            case "Double":
                                                excelData.headers[i].TypeValue = "Number";
                                                if (reader.IsDBNull(i) == true)
                                                {
                                                    record.Add(excelData.headers[i].Field, null);
                                                }
                                                else record.Add(excelData.headers[i].Field, reader.GetDouble(i));
                                                break;
                                          
                                        }
                                       
                                    }
                                }
                            }
                        if (row != 0) {
                            excelData.lstRecords.Add(record as ExpandoObject);   
                        }
                            row++;
                        }
                    
                }
            }
            return excelData;
        }
        static void Main(string[] args)
        {
            var excelLoadData = getRecordFromPath("data.xlsx");
        }
    }
}
