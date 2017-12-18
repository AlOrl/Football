using Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BL
{
   public  class DataProcessing:IDataProvider
    {
        public IDataProvider data;
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
