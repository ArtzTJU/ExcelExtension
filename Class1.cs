using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelExtension
{
    public static class Class1
    {
        public static string FoodPath
        {
            get
            {
                return Path.GetDirectoryName(ExcelDnaUtil.XllPath) + "\\Jidlo.csv";
            }
        }

        [ExcelFunction(Description = "Bilkoviny")]
        public static string Bilkoviny(string name, int grams)
        {
            try
            {
                var data = GetData();
                
                if (data.ContainsKey(name))
                    return (grams * data[name].Bilkoviny).ToString();
                else
                    return "Neni zapsano v jidlech";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Sacharidy")]
        public static string Sacharidy(string name, int grams)
        {
            try
            {
                var data = GetData();
                
                if (data.ContainsKey(name))
                    return (grams * data[name].Sacharidy).ToString();
                else
                    return "Neni zapsano v jidlech";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Tuky")]
        public static string Tuky(string name, int grams)
        {
            try
            {
                var data = GetData();
            
                if (data.ContainsKey(name))
                    return (grams * data[name].Tuky).ToString();
                else
                    return "Neni zapsano v jidlech";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        private static Dictionary<string, Food> GetData()
        {
            var lines = File.ReadAllLines(FoodPath);

            var result = new List<Food>();

            foreach (var line in lines.Skip(1))
            {
                var values = line.Split(';');
                result.Add(new Food
                {
                    Name = values[0],
                    Bilkoviny = Convert.ToInt32(values[1]),
                    Sacharidy = Convert.ToInt32(values[2]),
                    Tuky = Convert.ToInt32(values[3]),
                });
            }

            return result.ToDictionary(key => key.Name, value => value);
        }

        public class Food
        {
            public string Name { get; set; }
            public int Bilkoviny { get; set; }
            public int Sacharidy { get; set; }
            public int Tuky { get; set; }
        }
    }
}
