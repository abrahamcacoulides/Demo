using System.Collections.Generic;
using System.Linq;
using Engine.Models;

namespace Engine.Factories
{
    public static class DWGsFactory
    {
        private static List<DWG> _listOfDWGs;

        public static List<DWG> CreateDWGsList(List<string> dwgList)
        {
            _listOfDWGs = new List<DWG>();

            foreach(string path in dwgList)
            {
                List<string> words = path.Split('\\').ToList();
                var item = words.Last();
                string job_num_index = item.Split('.').ToList().First();
                string job_num = job_num_index.Split('-').ToList().First();
                string page_num = job_num_index.Split('-').ToList().Last();

                _listOfDWGs.Add(new DWG(job_num,page_num,job_num_index,path));
            }

            return _listOfDWGs;
        }
    }
}
