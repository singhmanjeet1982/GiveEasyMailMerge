using System;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using Microsoft.Extensions.Options;
using Giveaway.Infra;
using System.Threading;

namespace Giveaway.Helper
{
    public  class WordMailMerge:IWordMailMerge
    {
        private static object oFalse;

        private readonly AppSettings appSettings;
        private readonly IEmailSendGrid email;

        public WordMailMerge(IOptions<AppSettings> _appSettings, IEmailSendGrid _email)
        {
            appSettings = _appSettings.Value;
            email = _email;
        }
        
        /// <summary>
        /// Create Merge document and send email
        /// </summary>
        /// <param name="templateFilePath"></param>
        /// <param name="templateValues"></param>
        /// <param name="valuesFilePath"></param>
        public string generateEmail(string templateFilePath, System.Data.DataTable templateValues, string valuesFilePath)
        {
            //OBJECT OF MISSING "NULL VALUE"
            object oMissing = System.Reflection.Missing.Value;
            object oTemplatePath = templateFilePath;
            Application wordApp;
            Document wordDoc = new Document();


            try
            {

                foreach (DataRow dr in templateValues.Rows)
                {
                    wordApp = new Application();
                    wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
                    foreach (Field myMergeField in wordDoc.Fields)
                    {
                        Range rngFieldCode = myMergeField.Code;
                        String fieldText = rngFieldCode.Text;

                        // ONLY GETTING THE MAILMERGE FIELDS
                        if (fieldText.StartsWith(" MERGEFIELD"))
                        {
                            //logic to get the exact field name, we have to trim and extract fieldname
                            Int32 endMerge = fieldText.IndexOf("\\");
                            Int32 fieldNameLength = fieldText.Length - endMerge;
                            String fieldName = fieldText.Substring(11, endMerge - 11);
                            fieldName = fieldName.Trim();


                            //Fill the field with values from csv
                            if (templateValues.Columns.Contains(fieldName))
                            {
                                myMergeField.Select();
                                //Set default value here for any field
                                if (fieldName == "shippingcost" && dr[fieldName].ToString().Trim() == "")
                                {
                                    wordApp.Selection.TypeText("AUD 0.00");
                                }
                                wordApp.Selection.TypeText(dr[fieldName].ToString());

                            }
                            else
                            {
                                //Logic if field name is not there in the csv or set any default value here
                                myMergeField.Select();
                                wordApp.Selection.TypeText("Error Value");


                            }

                        }
                        
                    }


                    //Save word file as HTML to 
                    wordDoc.SaveAs(Path.GetDirectoryName(templateFilePath) + "\\myfile.html", WdSaveFormat.wdFormatHTML);
                    wordApp.Application.Quit();
                    wordDoc = null;
                    wordApp = null;

                    Thread.Sleep(1000);
                    //Read contents for HTML email and send email
                    string content = File.ReadAllText("upload\\myfile.html");
                    email.SendEmailAsync(dr["email"].ToString(), dr["subject"].ToString(), content).GetAwaiter().GetResult();



                }
                return "true";
            }
            catch (Exception ex)
            {
                // Release References.
                wordDoc = null;
                wordApp = null;
                if (ex.Source.ToLower().Contains("mail"))
                {
                    return "Error: Problem in sending email. Please check SMTP connection.";
                }
                throw;
            }
            finally
            {
                File.Delete("upload\\myfile.html");

            }

        }
    }
}
