using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace IntegradorWebService.Rest
{
    class RestPerfilImportacao
    {

        public static Rootobject ProcessaListaPerfil(string txtUsr, string txtPwd)
        {
            Rootobject b = new Rootobject();

            try
            {
                ServicePointManager.Expect100Continue = false;


                HttpWebRequest httpWReq = (HttpWebRequest)WebRequest.Create(IntegradorWebService.Properties.Settings.Default.Site + "/vipp/remoto/ListarPerfisRemoto.php");

                ASCIIEncoding encoding = new ASCIIEncoding();

                string postData = "";
                postData += "&Usr=" + txtUsr;
                postData += "&Pwd=" + txtPwd;

                byte[] data = encoding.GetBytes(postData);


                httpWReq.Method = "POST";
                httpWReq.ContentType = "application/x-www-form-urlencoded";
                httpWReq.ContentLength = data.Length;

                using (Stream stream = httpWReq.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                HttpWebResponse response = (HttpWebResponse)httpWReq.GetResponse();

                string responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();


                b = Newtonsoft.Json.JsonConvert.DeserializeObject<Rootobject>(responseString);


            }
            catch (Exception)
            {
                b.Data[0].NomePerfil = "Erro";

            }

            return b;
        }

    }
}
