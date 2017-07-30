using ExcelDataReader;
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
    class Program
    {

       
        static void Main(string[] args)
        {
            Console.Write("Start preprocess " + args[0]);
            var processCommand = args[0];
            ExcelData excelLoadData;
            switch (processCommand) {
                case "remove": // How to use: ProgramName remove field1,field2 C:\Input.xlsx D:\Output.xlsx
                    excelLoadData = ExcelService.getRecordFromPath(args[2]);
                    //Process data excel
                    var headersWillBeRemoved = args[1].Split(',');
                    excelLoadData.removeFields(headersWillBeRemoved);
                    ExcelService.ExportExcelInPath(args[3], excelLoadData);
                    break;
                case "standardized":
                    string typeStandardized = args[1];
                    switch (typeStandardized) {
                        case "min-max":
                            excelLoadData = ExcelService.getRecordFromPath(args[3]);
                            var headersWillBeStandardized = args[2].Split(',');
                            excelLoadData.StandardizedByMinMax(headersWillBeStandardized);
                            ExcelService.ExportExcelInPath(args[4], excelLoadData);
                            break;
                        case "z-score":
                            excelLoadData = ExcelService.getRecordFromPath(args[3]);
                            var headersWillBeStandardized1 = args[2].Split(',');
                            excelLoadData.StandardizedByZScore(headersWillBeStandardized1);
                            ExcelService.ExportExcelInPath(args[4], excelLoadData);
                            break;
                    }
                    break;
                case "binning":
                    string typeBinning = args[1];
                    switch (typeBinning)
                    {
                        case "width":
                            excelLoadData = ExcelService.getRecordFromPath(args[4]);
                            var _fieldWillBeBinning = args[2].Split(',');
                            excelLoadData.BinningWidth(_fieldWillBeBinning, int.Parse(args[3]));
                            ExcelService.ExportExcelInPath(args[5], excelLoadData);
                            break;

                            break;
                        case "height":
                            excelLoadData = ExcelService.getRecordFromPath(args[4]);
                            var fieldWillBeBinning = args[2].Split(',');
                            excelLoadData.BinningHeight(fieldWillBeBinning, int.Parse(args[3]));
                            ExcelService.ExportExcelInPath(args[5], excelLoadData);
                            break;
                    }
                    break;
                case "removeMissingInstance": //How to user ProgramName removeMissingInstance C:\Input.xlsx D:\Output.xlsx
                    excelLoadData = ExcelService.getRecordFromPath(args[2]);
                    excelLoadData.removeMissingInstance(args[1].Split(','));
                    ExcelService.ExportExcelInPath(args[3], excelLoadData);
                    break;  
            }
            Console.WriteLine("Done!"); 
        }
    }
}
