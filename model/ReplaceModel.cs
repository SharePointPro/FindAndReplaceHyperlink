using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplaceHyperlink.model
{
    public class ReplaceModel
    {
        public string From { get; set; }

        public string To { get; set; }


        /// <summary>
        /// Convert comma seperated string list into List of ReplaceModel 
        /// </summary>
        /// <param name="targetString"></param>
        /// <returns></returns>
        public static List<ReplaceModel> Convert(IEnumerable<string> targetString)
        {
            var replaceModelList = new List<ReplaceModel>();
            foreach (string str in targetString)
            {
                ReplaceModel replaceModel = new ReplaceModel();
                string[] strArray = str.Split(',');
                replaceModel.From = strArray[0].Trim();
                replaceModel.To = strArray[1].Trim();
                replaceModelList.Add(replaceModel);
            }
            return replaceModelList;
        }
    }
}
