using System;
using System.Collections.Generic;
using System.Globalization;
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
    public static class Extension
    {
        public static string FoodPath
        {
            get
            {
                return Path.GetDirectoryName(ExcelDnaUtil.XllPath) + "\\Jidlo.csv";
            }
        }

        [ExcelFunction(Description = "Proteins")]
        public static string Bilkoviny(string name, int grams)
        {
            try
            {
                var data = GetData();
                
                if (data.ContainsKey(name))
                    return (grams * data[name].Proteins).ToString();
                else
                    return "Neni zapsano v jidlech";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Carbohydrates")]
        public static string Sacharidy(string name, int grams)
        {
            try
            {
                var data = GetData();
                
                if (data.ContainsKey(name))
                    return (grams * data[name].Carbohydrates).ToString();
                else
                    return "Neni zapsano v jidlech";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Fats")]
        public static string Tuky(string name, int grams)
        {
            try
            {
                var data = GetData();
            
                if (data.ContainsKey(name))
                    return (grams * data[name].Fats).ToString();
                else
                    return "Neni zapsano v jidlech";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Calories")]
        public static string Kalorie(string name, int grams)
        {
            try
            {
                var data = GetData();

                if (data.ContainsKey(name))
                    return (grams * data[name].Calories).ToString();
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
                    Name = values[0].ToString(),
                    Proteins = double.Parse(values[1], NumberStyles.Any, CultureInfo.InvariantCulture),
                    Carbohydrates = double.Parse(values[2], NumberStyles.Any, CultureInfo.InvariantCulture),
                    Fats = double.Parse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture),
                    Calories = double.Parse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture),
                });
            }

            return result.ToDictionary(key => key.Name, value => value);
        }

        public class Food
        {
            public string Name { get; set; }
            public double Proteins { get; set; }
            public double Carbohydrates { get; set; }
            public double Fats { get; set; }
            public double Calories { get; set; }
        }
    }
}
