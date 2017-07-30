using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bai1B
{
    class Program
    {
        static void Main(string[] args)
        {
            CountryManager countryManager = new CountryManager();
            Console.WriteLine("Load from file...");
            countryManager.loadFromFilePath(args[0]);
            Console.WriteLine("Remove empty records...");
            countryManager.removeEmptyRecords();
            Console.WriteLine("Remove duplicate records...");
            countryManager.removeDuplicateInfo();
            Console.WriteLine("Export to file path...");
            countryManager.exportToFilePath(args[1]);
            Console.WriteLine("Remove records if field missing...");
            Process myProcess = new Process();
            var currentPath = Directory.GetCurrentDirectory();
            try
            {
                myProcess.StartInfo.UseShellExecute = false;
                // You can start any process, HelloWorld is a do-nothing example.
                
                Console.WriteLine("Execute " + String.Format(currentPath + "\\6_1A.exe removeMissingInstance area {0} {1}", args[1], args[1]));
                myProcess.StartInfo.FileName = String.Format("6_1A.exe");
                myProcess.StartInfo.Arguments = string.Format("removeMissingInstance area {0} {1}", args[1], args[1]);
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.Start();
                myProcess.WaitForExit();
                Console.WriteLine("Coppy xlsx to current execute...");
                File.Copy(args[1], currentPath + @"\6_1B.xlsx", true);
                Console.WriteLine("Done!");

            }
            
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            




        }
    }
}
