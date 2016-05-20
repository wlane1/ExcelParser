using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using FileHelpers;

namespace ExcelParser
{
    public class Program
    {
        private static readonly Dictionary<string, string> BondLookupList = new Dictionary<string, string>();
        private static List<Bond> BondsList = new List<Bond>();

        [STAThread]
        private static void Main(string[] args)
        {
            BuildBondList();
            UpdateExcelSheet();
            GenerateFinalCsv();

            Console.WriteLine("\nPress a key to exit ..");
            Console.ReadLine();
        }

        private static void BuildBondList()
        {
            const string file = "data.csv";
            string number = null;
            var readName = false;

            // iterate through the csv file
            var lines = File.ReadAllLines(file);
            foreach (var line in lines)
            {
                if (line.Contains(",Bond Number:,,"))
                {
                    var records = line.Split(',');
                    number = records[3];
                }

                if (line.Contains(",Partner Account Name,"))
                {
                    //rowString.Replace('"', '');
                    readName = true;
                    continue;
                }

                if (readName)
                {
                    var records = line.Split(',');
                    var name = records[7].Replace('"', ' ').Trim();

                    AddToList(number, name);

                    number = null;
                    readName = false;
                }
            }
        }

        private static void UpdateExcelSheet()
        {
            bool isFirstLine = true;
            const string file = "output-tab.txt";

            var lines = File.ReadAllLines(file);
            foreach (var line in lines)
            {
                if (isFirstLine)
                {
                    isFirstLine = false;
                    continue;
                }

                var records = line.Split('\t');

                if (records[0] == "9914B3963")
                {
                    continue;
                }

                var bond = new Bond();
                bond.BondNumber = records[0];
                bond.PrincipalName = BondLookupList[bond.BondNumber];
                bond.InputDate = records[2];
                bond.ImporterNumber = records[3];
                bond.Type = records[4];
                bond.EffMonth = records[5];
                bond.EffectiveDate = records[6];
                bond.BondAmount = records[7];
                bond.Address1 = records[8];
                bond.Address2 = records[9];
                bond.CityStateZip = records[10];
                bond.SuretyAgentId = records[11];

                BondsList.Add(bond);
            }
        }

        private static void GenerateFinalCsv()
        {
            var file = @"output.final.csv";

            using (StreamWriter sw = new StreamWriter(file))
            {
                int count = 1;
                foreach (var obj in BondsList)
                {
                    sw.WriteLine("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", obj.BondNumber, obj.PrincipalName, obj.InputDate, obj.ImporterNumber, obj.Type, obj.EffMonth,
                        obj.EffectiveDate, obj.BondAmount, obj.Address1, obj.Address2, obj.CityStateZip, obj.SuretyAgentId);
                    Console.WriteLine(count + " {0},{1}", obj.BondNumber, obj.PrincipalName + "\n");
                    count++;
                }
            }
        }

        private static void AddToList(string num, string name)
        {
            if (!string.IsNullOrEmpty(num) && !string.IsNullOrEmpty(name))
            {
                BondLookupList.Add(num, name);
            }
        }
    }

    public class Bond
    {
        public string BondNumber;
        public string PrincipalName;
        public string InputDate;
        public string ImporterNumber;
        public string Type;
        public string EffMonth;
        public string EffectiveDate;
        public string BondAmount;
        public string Address1;
        public string Address2;
        public string CityStateZip;
        public string SuretyAgentId;

    }
}
