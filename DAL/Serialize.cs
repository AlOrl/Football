using Contract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace BL
{
    public class Serialize : DataProvider
    {
        List<Question> list;
        public List<Question> GetQuestions()
        {
            try
            {
                XmlSerializer ser = new XmlSerializer(typeof(List<Question>));
                using (FileStream fs = new FileStream("kek.xml", FileMode.Open))
                {
                    list = ser.Deserialize(fs) as List<Question>;
                }
                return list;
            }
            catch(FileNotFoundException)
            {
                throw new FileNotFoundException();
            }
        }
    }
}
