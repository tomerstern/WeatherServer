using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Text;
using System.Web.Http;
using System.Web.Http.Cors;

namespace OwinSelfhost
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class ValuesController : ApiController
    {
        [HttpPost]
        public string GetWeatherByCity([FromBody]string city)
        {
            var location = string.Format("&q={0}", city);
            return GetWeather(location);
        }

        [HttpPost]
        public string GetWeatherByCoordinates([FromBody]string[] coordinates)
        {
            var location = string.Format("&lat={0}&lon={1}", coordinates[0], coordinates[1]);
            return GetWeather(location);
        }

        private string GetWeather(string location)
        {            
            var weatherApiUrl = ConfigurationManager.AppSettings["WeatherApiUrl"];
            var appId = ConfigurationManager.AppSettings["AppId"];
            var units = ConfigurationManager.AppSettings["Units"];
            var urlRequest = new StringBuilder(weatherApiUrl);
            urlRequest.AppendFormat("?units={0}", units);
            urlRequest.AppendFormat("&appid={0}{1}", appId, location);
            string result = string.Empty;
            try
            {
                result = (new WebClient()).DownloadString(urlRequest.ToString());
            }
            catch
            {
                return string.Empty;
            }
            return result.ToString();
        }
    }
}