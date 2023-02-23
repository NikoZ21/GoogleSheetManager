using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetApi
{

    /// <summary>
    /// TURN THE SPREAD SHEET INTO TEXT FILE TO COMMAND STANARD
    /// </summary>
    public class CommandsFileCreator
    {
        public static string filePath = string.Empty;

        /// <summary>
        /// WRITTE EACH ROW OF SPREAD SHEET INTO TEXT FILE
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="headers"></param>
        /// <param name="sheetName"></param>
        public static void CreateCommandFile(Dictionary<string, Dictionary<string, Dictionary<string, string>>> delta, List<Sheet> sheets)
        {
            if (delta == null || delta.Count == 0)
            {
                return;
            }

            foreach (var sheet in sheets)
            {
                var tab = delta[sheet.GetName()];

                foreach (var key in sheet.GetRowKeys())
                {
                    if (!tab.ContainsKey(key))
                    {
                        Console.WriteLine("here");
                        continue;
                    }

                    var fileName = SpreadSheet.Name + sheet.GetName() + key;

                    using (StreamWriter writer = new StreamWriter(filePath + $"/{fileName}.txt"))
                    {
                        string result = string.Empty;

                        foreach (var header in sheet.GetHeaders())
                        {
                            if (header == "/")
                            {
                                result += header + ":" + "" + ";";
                            }
                            else
                            {
                                result += header + ":" + tab[key][header] + ";";
                            }
                        }

                        writer.WriteLine(result);
                    }
                }

            }

        }
    }
}
