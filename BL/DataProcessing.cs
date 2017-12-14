using Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BL
{
   public  class DataProcessing:DataProvider
    {
        public DataProvider data;
        public DataProcessing()
        {
            data = new Serialize();
        }
       public List<Question> GetQuestions()
        {
            return data.GetQuestions();
        }
    }
}
