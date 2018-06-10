using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Collections.Generic;
namespace JsonUtilityLib
{
    public class JsonUtility
    {
        public static async Task<T[]> GetAllObject<T>(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var tmp = client.GetAsync(url).Result;
                List<T> arr = new List<T>();
                if (!tmp.IsSuccessStatusCode)
                {
                    throw new Exception("return code is not suceed!");
                }
                else
                {
                    var result = await tmp.Content.ReadAsStringAsync();
                    if (result.TrimStart ().StartsWith ("["))
                    {
                        JArray a = JArray.Parse(result);
                        var ch = a.Children<JObject>(); 
                        foreach (JObject j in ch)
                        {
                            arr.Add(j.ToObject<T>());
                        }
                    }
                    else
                    {
                        var t = JObject.Parse(result).ToObject<T>();
                        arr.Add(t);
                    }
                }
                return arr.ToArray ();
            }
        }
    }
}
