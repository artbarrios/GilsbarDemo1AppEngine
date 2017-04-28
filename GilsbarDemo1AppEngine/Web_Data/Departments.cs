using GilsbarDemo1.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace GilsbarDemo1AppEngine.Web_Data
{
    class DepartmentsWebData
    {
        // global static vars
        private static HttpClient client = new HttpClient();

        // GET: api/DepartmentsData
        public static List<Department> GetDepartments()
        {
            // return the data or perform an action using the remote webApiUrl
            string webApiPath = "api/DepartmentsData";
            string results = "";
            try
            {
                results = client.GetAsync(AppCommon.BuildUrl(AppCommon.GetRemoteWebApiUrl(), webApiPath)).Result.Content.ReadAsStringAsync().Result;
                return JsonConvert.DeserializeObject<List<Department>>(results);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("GetDepartments: " + e.Message, e);
                throw new Exception(message);
            }
        } // GetDepartments
    }
}

