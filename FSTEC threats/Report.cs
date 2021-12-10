using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FSTEC_threats
{
    class Report
    {
        List<string> newThreatList = new List<string>();
        List<string> changesList = new List<string>();

        public void AddChange(string id, string field, string old, string current)
        {
            changesList.Add($"Изменена УБИ.{id}\nПоле '{field}' было: \n{old}\nCтало:\n{current}\n");
        }

        public void AddNewThreat(string id, string name)
        {
            changesList.Add($"Добавлена новая УБИ.{id} - {name}\n");
        }

        public override string ToString()
        {
            if (newThreatList.Count == 0 && changesList.Count == 0)
                return "Никаких изменений не произошло";
            string res = "";
            foreach (var newThreat in newThreatList)
            {
                res += newThreat + "\n";
            }
            foreach (var change in changesList)
            {
                res += change + "\n";
            }
            return res;
        }
    }
}
