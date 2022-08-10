using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CRM_Publisher_V2
{
    class ExceptionLogging
    {
        private static String Errormsg;

        public static void SendErrorToText(string ex)
        {
            var line = Environment.NewLine + Environment.NewLine;

            Errormsg = ex.ToString();


            try
            {
                string currentpath = Directory.GetCurrentDirectory();
                //string filepath = currentpath + "/ExceptionDetailsFiles/";
                string filepath = "C:/ABL_CRM_Publisher/Log_Failure//";
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
                    string error = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Error Message:" + " " + Errormsg + line;
                    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
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
