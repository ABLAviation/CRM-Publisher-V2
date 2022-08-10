using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using TextBox = Microsoft.Office.Interop.Excel.TextBox;
using Microsoft.Xrm.Tooling.Connector;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Configuration;
using Microsoft.Xrm.Sdk.Query;
using System.Text.RegularExpressions;
using System.Activities.Statements;
using System.Windows.Navigation;
using System.Data;
using System.Reflection;
using System.Net;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace CRM_Publisher_V2
{
    public partial class CRM_Pub
    {
        private void CRM_Pub_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnUpdateCRM_Click(object sender, RibbonControlEventArgs e)
        {
            //Calling the "Parameters" and "Update CRM" sheets and necessary Ranges
            Worksheet ws = Globals.ThisAddIn.GetUpdCRMSheet();
            Worksheet Paramsws = Globals.ThisAddIn.GetParamWorksheet();

            //string Publish_Flag = (string)((Range)ws.Application.get_Range("Publish_Flag")).Value2;
            bool Publish_Flag_Value = ws.Application.get_Range("Publish_Flag").Value2;
            string Publish_Flag = Publish_Flag_Value.ToString();

            if (Publish_Flag != null && Publish_Flag.ToUpper() == "TRUE")
            {
                

                try
                {

                    //Initiating Username and Password form
                    Form1 form = new Form1();
                    form.ShowDialog();

                    //Declaring different necessary variables 
                    string param2_username = form.Email;
                    string param3_password = form.Password;

                    MessageBox.Show(param2_username.Substring(param2_username.Length - 16, 16));

                    string param4_entity = "";
                    string param5_field = "";
                    string param8_fieldtype = "";
                    string param6_data = "";
                    string param7_rowid = "";

                    // In case user wants to cancel
                    if (string.IsNullOrEmpty(param2_username) || string.IsNullOrEmpty(param3_password))
                    {

                        goto End;
                    }

                
                    //Range rng = (Range)ws.Application.get_Range("range");
                    Range Params = (Range)ws.Application.get_Range("CRMOutputData");
                    //string RecorID = (string)Params[1, 5];

                  /*  int totalRows = Params.Rows.Count;
                    int totalColumns = Params.Columns.Count;
                    for (int rowCounter = 1; rowCounter <= totalRows; rowCounter++)
                    {
                        for (int colCounter = 1; colCounter <= totalColumns; colCounter++)
                        {
                            var cellVal = ws.Cells[rowCounter+4, colCounter+8];
                            var val = cellVal.Text;
                            MessageBox.Show(val);
                        }
                    }*/



                    var array = Params.Value;
                    int x = 0;
                    

                    //Get last non empty row in range
                    foreach (Range row in Params.Rows)
                    {
                        var check = row.Cells[1, 6].Value2;
                        if (!string.IsNullOrEmpty(check))
                        {
                            x += 1;
                        }


                    }

                    for( var i = 1; i<=x ; i++)
                    {

                        var FieldValue = Params[i, 2].Value2;
                        var Field = Params[i, 1].Value2;

                        // MessageBox.Show(DataType.GetType().Name+" -------- "+Field);
                        switch (Field)
                        {
                            case "Life at Ext. 1 (%)":
                                if (FieldValue.GetType().Name != "Double")
                                {
                                    MessageBox.Show("Please check \"" + Field + "\" Value !");
                                    
                                    goto End;
                                };
                                break;
                            case "MR Balance at Ext. 1 ($)":
                                if (FieldValue.GetType().Name != "Double")
                                {
                                    MessageBox.Show("Please check \"" + Field + "\" Value !");
                                    
                                    goto End;
                                };
                                break;
                            case "Enable on FF":
                                if (FieldValue.ToUpper() != "NO" && FieldValue.ToUpper() != "YES")
                                {
                                    MessageBox.Show("Please check \"" + Field + "\" Value !");
                                    goto End;
                                };
                                break;
                            case "Internal Notes":
                                if (FieldValue.GetType().Name != "String")
                                {
                                    MessageBox.Show("Please check \"" + Field + "\" Value !");
                                    
                                    goto End;
                                };
                                break;
                            case "Conclusions":
                                if (FieldValue.GetType().Name != "String")
                                {
                                    MessageBox.Show("Please check \"" + Field + "\" Value !");
                                    
                                    goto End;
                                };
                                break;
                            case "Engine Conclusions":
                                if (FieldValue.GetType().Name != "String")
                                {
                                    MessageBox.Show("Please check \"" + Field + "\" Value !");
                                    
                                    goto End;
                                };
                                break;
                            default:
                                break;

                        }



                    }

                    //string ParamsVal = Params.Value2;
                    Range URLRng = (Range)ws.Application.get_Range("CRM_URL");
                    string param1_url = URLRng.Value2;

                    String ParamList;
                    List<string> ParamListArray = new List<string>();
                    foreach (Range row in Params.Rows)
                    {
                        var check = row.Cells[1, 6].Value2;
                        if (!string.IsNullOrEmpty(check))
                        {
                            ParamList = ("\"Entity Name\"::\"" + row.Cells[1, 5].Value2.ToString() + "\"//\"Entity Field\"::\"" + row.Cells[1, 3].Value2.ToString() + "\"//\"Field Type\"::\"" + row.Cells[1, 4].Value2.ToString() + "\"//\"Data value to send\"::\"" + row.Cells[1, 2].Value2.ToString() + "\"//\"Record ID\"::\"" + row.Cells[1, 6].Value2.ToString() + "\"");


                            ParamListArray.Add(ParamList);
                            //MessageBox.Show(ParamList);

                        }


                    }

                    string ParamsVal = String.Join("||", ParamListArray);
                    //ws.Range["O1"].Value = ParamsVal;
                    //MessageBox.Show(ParamsVal);
                    //Connecting Excel to D365
                    string connectionstring_dynamic = "AuthType=OAuth;Username=" + param2_username + ";Password=" + param3_password + ";Url = " + param1_url + ";AppId=fb6c6ce9-b188-4517-a971-02fe33362a16;RedirectUri=https://crmpublisher-console;LoginPrompt=auto;";

                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var servicedynamic = new CrmServiceClient(connectionstring_dynamic);

                    WhoAmIRequest systemUserRequest = new WhoAmIRequest();
                    WhoAmIResponse systemUserResponse = (WhoAmIResponse)servicedynamic.Execute(systemUserRequest);
                    Guid userId = systemUserResponse.UserId;

                    // Lookup User
                    var User = servicedynamic.Retrieve("systemuser", userId, new ColumnSet("fullname"));
                    string fullName = User["fullname"].ToString();

                    

                    /*//Informing the User whether the connection is successful
                    if (servicedynamic.IsReady)
                    {
                       //MessageBox.Show("Connection succeeded !");
                    
                    }*/

                    // Splitting the Params string from Excel (Parameters sheet) and storing their values in the previously declared variables
                    //string[] OutputsParams = ParamsVal.Split(new[] { "||" }, StringSplitOptions.None);
                    //MessageBox.Show(OutputsParams[0]);
                    for (int j = 0; j < ParamListArray.Count; j++)
                    {
                        string[] LineParams = ParamListArray[j].Split(new[] { "//" }, StringSplitOptions.None);

                        for (int k = 0; k < LineParams.Length; k++)
                        {
                            string[] strall;
                            strall = LineParams[k].Split(new[] { "::" }, StringSplitOptions.None);
                            //MessageBox.Show(strall[0] + " = " + strall[1]);
                            switch (strall[0])
                            {
                                case @"""Entity Name""":
                                    param4_entity = strall[1].Replace("\"", string.Empty).Trim();
                                    break;
                                case @"""Entity Field""":
                                    param5_field = strall[1].Replace("\"", string.Empty).Trim();
                                    break;
                                case @"""Field Type""":
                                    param8_fieldtype = strall[1].Replace("\"", string.Empty).Trim();
                                    break;
                                case @"""Data value to send""":
                                    param6_data = strall[1].Replace("\"", string.Empty).Trim();
                                    break;
                                case @"""Record ID""":
                                    param7_rowid = strall[1].Replace("\"", string.Empty).Trim();
                                    break;

                            }
                        }
                        //Retrieving the Org Name
                        string orgname = servicedynamic.ConnectedOrgUniqueName;
                    
                        //Calling the method updating the record
                        if (orgname != "" && orgname != null)
                        {
                            //MessageBox.Show(param7_rowid);
                            updaterecord_FromGUID_dynamicvalues(servicedynamic, param4_entity, param5_field, param8_fieldtype, param6_data, param7_rowid, param1_url, param2_username);
                            ws.get_Range("Update_Status").Value2 = "Upate Successful !";

                            ws.get_Range("Updated_On").Value2 = System.DateTime.Now.ToString("yyyy'/'MM'/'dd'  'HH':'mm':'ss");
                            ws.get_Range("Updated_By").Value2 = fullName;

                        }
                        else
                        {
                            //Informing the User in case of failed connection
                            //MessageBox.Show(orgname);
                            ws.get_Range("Update_Status").Value2 = ("Couldn't retrieve OrgName for record ID : " +param7_rowid+" \nPlease contact IT!");
                            
                            ExceptionLogging.SendErrorToText("Connection could not be established !");
                            
                        }

                    }
                    MessageBox.Show("Program Terminated!");
                End:
                    if (string.IsNullOrEmpty(param2_username) || string.IsNullOrEmpty(param3_password))
                    {
                        MessageBox.Show("Operation Canceled !");
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    ExceptionLogging.SendErrorToText(ex.Message);
                    // Console.ReadLine();
                }
            }
            else
            {
                MessageBox.Show("Please check Publish Flag value !!! ");
            }
        }

        //Record updating Method
        static void updaterecord_FromGUID_dynamicvalues(CrmServiceClient args, string entityname, string fieldname, string fieldtype, string fieldvalue, string recordid, string crm_url, string crm_user)
        {
            decimal fieldvaluenum = 0;
            string fieldvaluestring = null;
            Int32 fieldvalueint = 0;
            string crm_field_type = "";

            try

            {

                string stringGuid = recordid;

                //Parsing the Record ID passed from the Main Method
                Guid IDGuid = Guid.Parse(stringGuid);

                //Initiating the Entity
                Entity entitytoupdate = new Entity(entityname);

                //Retrieving Entity's Metadata (String values)
                entitytoupdate = args.Retrieve(entityname, IDGuid, new ColumnSet(true));

                RetrieveEntityRequest req = new RetrieveEntityRequest();

                req.RetrieveAsIfPublished = true;

                req.LogicalName = entityname;

                req.EntityFilters = EntityFilters.Attributes;

                RetrieveEntityResponse resp = (RetrieveEntityResponse)args.Execute(req);
                EntityMetadata em = resp.EntityMetadata;
                /*MessageBox.Show("hello "+em.LogicalName);*/
                foreach (AttributeMetadata a in em.Attributes)
                {
                    if (a.LogicalName == fieldname)
                    {

                        //MessageBox.Show(a.LogicalName+"------"+ fieldname); //"Decimal" "Integer"
                        crm_field_type = a.AttributeType.ToString();
                        //MessageBox.Show(crm_field_type);
                        break;
                    }

                }
                
                //Parsing Excel Params to their corresponding CRM data types
                switch (crm_field_type)
                {
                    case "Decimal":
                        fieldvaluenum = decimal.Parse(fieldvalue);
                        /**//*MessageBox.Show(fieldvaluenum+"");*/
                        break;
                    case "Integer":
                        fieldvalueint = Int32.Parse(fieldvalue);
                        break;
                    default:
                        fieldvaluestring = fieldvalue;
                        break;
                }

                //Option Sets are a special case : we need a new method to assign the new values (retrieveoptionsetvaluefromtext)

                string caseSwitch = fieldtype;
                
                switch (caseSwitch)
                {
                    case "optionset":
                        int my_optionsetvalue = retrieveoptionsetvaluefromtext(args, entityname, fieldname, fieldvalue);

                        entitytoupdate[fieldname] = new OptionSetValue(my_optionsetvalue);
                        //MessageBox.Show(entitytoupdate[fieldname] + "");
                        break;

                    //Params string only differentiate between "Simplefield" and "Optionset"  -> we need to parse the "Simplefield" input values into their right ouput data type
                    case "simplefield":

                        switch (crm_field_type)
                        {
                            case "Decimal":
                                entitytoupdate[fieldname] = fieldvaluenum;

                                break;
                            case "Integer":
                                entitytoupdate[fieldname] = fieldvalueint;
                                break;
                            default:
                                entitytoupdate[fieldname] = fieldvaluestring;
                                break;
                        }
                        
                        break;
                    default:

                        break;
                }
                //MessageBox.Show(fieldname+"------"+ entitytoupdate[fieldname]);
                //Updating the record
                //MessageBox.Show(entitytoupdate[fieldname]+"");
                args.Update(entitytoupdate);

                /*Worksheet ws = Globals.ThisAddIn.GetUpdCRMSheet();
                */

                //Filling the success log file                
                SucceededLogging.SendNotifToText(entityname, fieldname, fieldvalue, crm_url, crm_user);

            }



            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                ExceptionLogging.SendErrorToText(e.Message);

            }

        }


        static int retrieveoptionsetvaluefromtext(CrmServiceClient args, string entityname, string fieldname, string fieldvalue)
        {
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityname,
                LogicalName = fieldname,
                RetrieveAsIfPublished = true
            };

            var attributeResponse = (RetrieveAttributeResponse)args.Execute(attributeRequest);
            var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            var optionList = (from o in attributeMetadata.OptionSet.Options
                              select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();


            var activeValue = optionList.Where(o => o.Text == fieldvalue)
                                    .Select(o => o.Value)
                                    .FirstOrDefault();

            int my_optionsetvalue = (int)activeValue;


            return my_optionsetvalue;
        }


        


    }/////////

   
}

