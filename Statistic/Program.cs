using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Statistic
{
    class Program
    {
        public string data;
        public List<DataStatiscics> dataStatiscics;
        public string text;
        string pathCvs;
     

        public Program(string pathCsv)
        {
            dataStatiscics = new List<DataStatiscics>();
            Create(pathCsv);
            
        }

        void Create(string pathCsv)
        {
           
            // читаем csv файл и записываем в переменную text
            using StreamReader sr1 = new StreamReader(pathCsv, System.Text.CodePagesEncodingProvider.Instance.GetEncoding(1251));
            {
                text = sr1.ReadToEnd();
            }

            // с помощю регулярного выражения находим камеры КПА и создаем свой тип данных DataStatiscics для удобства и находим дату. Пилим это дела в Спиок

            Regex regex = new Regex(@"(([;]{1,10}\s+)(?<data>\d{2}\.\d{2}\.\d{4}))|(?<DataStatiscics>(?<border>KPA-\d+-\d+);КПА;;(?<detections>\d+))", RegexOptions.Multiline) ;
            MatchCollection matches = regex.Matches(text);
            
            foreach (Match m in matches)
            {
                GroupCollection groups = m.Groups;
                if (groups["DataStatiscics"].Success)
                {
                    dataStatiscics.Add(new DataStatiscics() {border = groups["border"].Value, detections = groups["detections"].Value });
                }
                else 
                {
                    if (groups["data"].Success)
                        data = groups["data"].Value;
                }
            }
        }

        static void Main(string[] args)
        {
            try
            {
                string[] filesCsv = Directory.GetFiles(Environment.CurrentDirectory + "\\csvFiles");
                foreach (string i in filesCsv)
                {
                    Program prog = new Program(i);
                    new Excel(prog.dataStatiscics, prog.data);
                }
                Console.Read();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.Read();
            }
         }
    }

}
