using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.IO;
using Newtonsoft.Json;
using System.Net;
using Contract;

namespace APIWeather
{
    public class WorkWithWeather
    {
        const string APPID = "8e534cdfe7d49b057a8f7d7b91b4451a";
        const int timeout = 5; //В секундах

        public string GetTemperature()
        {
            try
            {
                var request = new WebClient().DownloadString($"https://api.openweathermap.org/data/2.5/weather?lat=55&lon=37&APPID={APPID}");
                RootObject eoot = JsonConvert.DeserializeObject<RootObject>(request);
                string tr = ((int)eoot.main.temp - 273).ToString();
                return tr;
            }
            catch (WebException ex)
            {
                Logger.Log.Error($"Отсутствует подключение к интернету{ex}");
                throw new Exception();
            }
            catch (Exception ex)
            {
                Logger.Log.Error($"Неизвестная ошибка{ex}");
                throw new Exception();
            }



        }


    }
}
