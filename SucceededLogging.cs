using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using context = System.Web.HttpContext;

namespace CRM_Publisher_V2
{
    class SucceededLogging
    {
        private static String Infomsg_entityname, Infomsg_fieldname, Infomsg_fieldvalue, Infomsg_crmurl, Infomsg_crmuser;

        public static void SendNotifToText(string entityname, string fieldname, string fieldvalue, string crm_url, string crm_user)
        {
            var line = Environment.NewLine + Environment.NewLine;

            Infomsg_entityname = entityname.ToString();
            Infomsg_fieldname = fieldname.ToString();
            Infomsg_fieldvalue = fieldvalue.ToString();
            Infomsg_crmurl = crm_url.ToString();
            Infomsg_crmuser = crm_user.ToString();

            try
            {
                string currentpath = Directory.GetCurrentDirectory();
                //string filepath = currentpath + "/SucceededUpdatesDetailsFiles/";
                string filepath = "C:/ABL_CRM_Publisher/Log_Success/";
                //string filepath = context.Current.Server.MapPath("~/ExceptionDetailsFile/");  //Text File Path

                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);

                }
                filepath = filepath + DateTime.Today.ToString("dd-MM-yy") + ".txt";   //Text File Name
                if (!File.Exists(filepath))
                {


                    File.Create(filepath).Dispose();

                }
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    string error = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Notification Message:" + line + "CRM Environment >> " + Infomsg_crmurl + line + " Entity updated >> " + Infomsg_entityname + line + " field updated >> " + Infomsg_fieldname + line + " with new value >> " + Infomsg_fieldvalue + line + "Done by >> " + Infomsg_crmuser + line;
                    sw.WriteLine("-----------Notification Details on " + " " + DateTime.Now.ToString() + "-----------------");
                    sw.WriteLine("-------------------------------------------------------------------------------------");
                    sw.WriteLine(line);
                    sw.WriteLine(error);
                    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    sw.WriteLine(line);
                    sw.Flush();
                    sw.Close();

                }

            }
            catch (Exception e)
            {
                e.ToString();

            }
        }
    }
}

