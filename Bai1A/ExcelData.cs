using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bai1A
{
    class ExcelData
    {
        public List<ExpandoObject> lstRecords { get; set; }
        public List<TitleHeader> headers { get; set; }
        public ExcelData() {
            lstRecords = new List<ExpandoObject>();
            headers = new List<TitleHeader>();
        }
        public void removeFields(string[] fields) {
            foreach (String field in fields) {
                var header = headers.Find(hd => hd.Field == field);
                if (header != null) {
                    //Remove this header
                    headers.Remove(header);
                    //Remove data of Column header owner too
                    foreach (IDictionary<string, Object> record in lstRecords) {
                        record.Remove(header.Field);
                    }
                }
            }
        }

        public void removeMissingInstance(string[] fields)
        {
            List<ExpandoObject> lstRecordsWillBeRemoved = new List<ExpandoObject>();
            foreach (IDictionary<string, Object> record in lstRecords) {
                foreach (String headerName in fields) {
                    if (record[headerName] == null) {
                        lstRecordsWillBeRemoved.Add(record as ExpandoObject);
                        break;
                    }
                }
            }
            foreach (ExpandoObject record in lstRecordsWillBeRemoved) {
                lstRecords.Remove(record);
            }
        }
        private void StandardizedMinMaxByColumn(string field) {
            List<double> lstValue = new List<double>();
            foreach (IDictionary<string, object> record in lstRecords) {
                if (record[field] != null) {
                    lstValue.Add((double)record[field]);
                }
            }
            double min = lstValue.Min();
            double max = lstValue.Max();

            foreach (IDictionary<string, object> record in lstRecords)
            {
                if (record[field] != null)
                {
                    var valueRecord = (double)record[field];
                    record[field] = (valueRecord - min) / (max - min);
                }
            }

        }
        public void StandardizedByMinMax(string[] headersWillBeStandardized)
        {
            foreach (String header in headersWillBeStandardized) {
                var headerTitle = headers.Find(h => h.Field == header);
                if (headerTitle != null && headerTitle.TypeValue == "Number")
                {
                    StandardizedMinMaxByColumn(header);
                }
            }
        }
        private void StandardizedZScoreByColumn(string field)
        {
            List<double> lstValue = new List<double>();
            foreach (IDictionary<string, object> record in lstRecords)
            {
                if (record[field] != null)
                {
                    lstValue.Add((double)record[field]);
                }
            }
            double Sum = lstValue.Sum();
            int lenghtOfLst = lstValue.Count();
            double mean = Sum / lenghtOfLst;
            double temp = 0;
            foreach (double value in lstValue) {
                temp += (double)Math.Pow(value - mean, 2);
            }
            double std = (double)Math.Sqrt(temp / (lenghtOfLst - 1));



            foreach (IDictionary<string, object> record in lstRecords)
            {
                if (record[field] != null)
                {
                    var valueRecord = (double)record[field];
                    record[field] = (valueRecord - mean) / (std);
                }
            }

        }
        public void StandardizedByZScore(string[] headersWillBeStandardized)
        {
            foreach (String header in headersWillBeStandardized)
            {
                var headerTitle = headers.Find(h => h.Field == header);
                if (headerTitle != null && headerTitle.TypeValue == "Number")
                {
                    StandardizedZScoreByColumn(header);
                }
            }
        }
        public List<List<T>> Split<T>(List<T> array, int size)
        {
            List<List<T>> lst = new List<List<T>>();
            for (var i = 0; i < (float)array.Count / size; i++)
            {
                lst.Add(array.Skip(i * size).Take(size).ToList());
            }
            return lst;
        }
        public void BinningHeight(string field, int bin) {

            List<ExpandoObject> lstHavingValue = lstRecords.Where(h =>
             (h as IDictionary<string, object>)[field] != null)
                .OrderBy(h => (h as IDictionary<string, object>)[field]).ToList();

            var lst = Split<ExpandoObject>(lstHavingValue, bin).ToList();

            foreach (List<ExpandoObject> l in lst) {
                var lstNumber = l.Select(h => (Double)(h as IDictionary<string, Object>)[field]);
                var avg = lstNumber.Average();
                foreach (IDictionary<string, Object> i in l) {
                    i[field] = avg;
                }
            }

        }
        public void BinningHeight(string[] headersBinning, int bin)
        {
            foreach (String header in headersBinning)
            {
                var headerTitle = headers.Find(h => h.Field == header);
                if (headerTitle != null && headerTitle.TypeValue == "Number")
                {
                    BinningHeight(header, bin);
                }
            }
        }

        public List<BinningWidth> SplitWidthAndBinning(List<ExpandoObject> lstRecord, double min, double max, int bin, string field) {
            List<BinningWidth> lst = new List<BinningWidth>();

            while (min < max) {
                BinningWidth bw = new BinningWidth(min, min + bin);
                min += bin;
                bw.lstRecord = lstRecord.Where(h =>
                       (Double)(h as IDictionary<string, object>)[field] >= bw.Min
                    && (Double)(h as IDictionary<string, object>)[field] < bw.Max).ToList();
                if (bw.lstRecord.Count > 1) {
                    var avg = bw.lstRecord.Select(h => (Double)(h as IDictionary<string, Object>)[field]).Average();
                    foreach (IDictionary<string, object> record in bw.lstRecord) {
                        record[field] = avg;
                    }
                }
            }
            return lst;

        }

        public void BinningWidth(string field, int bin) {
            List<ExpandoObject> lstHavingValue = lstRecords.Where(h =>
            (h as IDictionary<string, object>)[field] != null)
               .OrderBy(h => (h as IDictionary<string, object>)[field]).ToList();

            double min = lstHavingValue.Select(h => (Double)(h as IDictionary<string, Object>)[field]).Min();
            double max = lstHavingValue.Select(h => (Double)(h as IDictionary<string, Object>)[field]).Max();
            SplitWidthAndBinning(lstHavingValue, min, max, bin, field);
            
        }
        public void BinningWidth(string[] headersBinning, int bin) {
            foreach (String header in headersBinning)
            {
                var headerTitle = headers.Find(h => h.Field == header);
                if (headerTitle != null && headerTitle.TypeValue == "Number")
                {
                    BinningWidth(header, bin);
                }
            }
        }


    }
}
