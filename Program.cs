using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using System.Security;
using System.Net;
using System.Linq;
using System.Collections.Generic;
using System.Configuration;

namespace JDETakeMail
{
    

    class Program
    {
        static void Main(string[] args)
        {
            Logger log = new Logger(String.Format("JDETakeMail_log_{0}", DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss")), @"D:\logs\JDETakeMail");
            try
            {
                log.WriteLine("START!");
                
                GetMail(log);

                log.WriteLine("Program successfully finished its work.");
            }
            catch (Exception ex)
            {
                log.WriteLine(ex.Message);
            }
        }

        private static void GetMail(Logger log)
        {
            string body, bodyHTML, sub, address = "", dw = "", ph = "", Gender = "", EmailAddr = "", Ag = "", CompName = "", 
                   MachineCode = "", Comment = "", Postcode = "", pib = "", region = "", whatType = "", typeOfComplaint = "",
                   datetimeSent, addressFrom, addressTo;
            int mailType = 0;
            

            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));//searching for new letters
            ItemView view = new ItemView(20);

            log.WriteLine("Trying to authenticate...");
            ExchangeService service = Authenticator.Authenticate(
                ConfigurationManager.AppSettings["login"], 
                ConfigurationManager.AppSettings["password"], 
                ConfigurationManager.AppSettings["domain"],
                ConfigurationManager.AppSettings["link"]);
            log.WriteLine("Authentication success");

            Mailbox mailBox = new Mailbox(ConfigurationManager.AppSettings["mail"]);
            FolderId folderProcessedId = GetExchangeFolderIdByName("Processed", service, mailBox);

            // Getting all letters from inbox
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);


            foreach (Item item in findResults.Items)
            {
                if (item is EmailMessage)
                {
                    PropertySet psPropsetHTML = new PropertySet
                    {
                        RequestedBodyType = BodyType.HTML,
                        BasePropertySet = BasePropertySet.FirstClassProperties
                    };
                    PropertySet psPropsetTxt = new PropertySet
                    {
                        RequestedBodyType = BodyType.Text,
                        BasePropertySet = BasePropertySet.FirstClassProperties
                    };


                    EmailMessage message = EmailMessage.Bind(service, item.Id, psPropsetTxt);
                    sub = message.Subject;

                    body = message.Body.Text;
                    datetimeSent = message.DateTimeSent.ToString("d MMMM yyyy г. HH:mm", new CultureInfo("uk-UA"));
                    log.WriteLine("Sent date: " + datetimeSent);
                    addressTo = message.DisplayTo;
                    log.WriteLine("Sent to: " + addressTo);

                    addressFrom = message.From.Address.ToString();

                    log.WriteLine("Sent from: " + addressFrom);

                    EmailMessage messageHTML = EmailMessage.Bind(service, item.Id, psPropsetHTML);
                    bodyHTML = messageHTML.Body.Text;

                    if (addressTo == null)
                        addressTo = "";
                    else
                        addressTo = addressTo.Replace("'", "''");
                    if (sub == null)
                        sub = "";
                    else
                        sub = sub.Replace("'", "''");

                    if (body == null)
                        body = "";
                    else
                        body = body.Replace("'", "''");

                    if (bodyHTML == null)
                        bodyHTML = "";
                    else
                        bodyHTML = bodyHTML.Replace("'", "''");

                    // Parsing taken letter
                    if (sub.Contains("Quote request online leads"))
                    {
                        if (body.Contains("Name"))
                            pib = TakeName(body, "Name");

                        if (body.Contains("Email"))
                            EmailAddr = TakeLineFromMail(body, "Email: ", 150);

                        if (body.Contains("PhoneNumber: "))
                            ph = TakePhone(body, "PhoneNumber: ");

                        if (body.Contains("CompanyName: "))
                            CompName = TakeLineFromMail(body, "CompanyName: ", 100);

                        if (body.Contains("Gender: "))
                            Gender = TakeLineFromMail(body, "Gender: ", 20);

                        if (body.Contains("Postcode: "))
                        {
                            TakeLineFromMail(body, "Postcode: ", 50);
                            Postcode = Regex.Replace(Postcode, @"[^\d]", "", RegexOptions.Compiled);
                        }
                        else
                            Postcode = "";

                        if (body.Contains("AdditionalComments: "))
                            Comment = TakeLineFromMail(body, "AdditionalComments: ", 500);

                        if (body.Contains("MachineCode: "))
                            MachineCode = TakeLineFromMail(body, "MachineCode: ", 50);
                        else
                            MachineCode = "";

                        mailType = 1;
                    }
                    else if (sub.Contains("Contact"))
                    {
                        if (body.Contains("Name"))
                            pib = TakeName(body, "Name");

                        if (body.Contains("Company name : "))
                            CompName = TakeLineFromMail(body, "Company name : ", 100);

                        if (body.Contains("E-mail : "))
                            EmailAddr = TakeLineFromMail(body, "E-mail : ", 150);

                        if (body.Contains("Phone Number : "))
                            ph = TakePhone(body, "Phone Number : ");

                        if (body.Contains("Salutation : "))
                            Gender = TakeLineFromMail(body, "Salutation : ", 20);
                        else
                            Gender = "";

                        if (body.Contains("Coffee machine : "))
                            MachineCode = TakeLineFromMail(body, "Coffee machine : ", 50);
                        else
                            MachineCode = "";

                        if (body.Contains("Adress : "))//it is intended; it is mistake in client`s template
                            address = TakeLineFromMail(body, "Adress : ", 250);
                        else
                            address = "";


                        if (body.Contains("Question : "))
                            dw = TakeLineFromMail(body, "Question : ", 100);
                        else
                            dw = "";

                        if (body.Contains("Your question : "))
                            Comment = TakeLineFromMail(body, "Your question : ", 500);

                        if (body.Contains("Agreement : "))
                            Ag = TakeLineFromMail(body, "Agreement : ", 50);
                        else
                            Ag = "";

                        mailType = 2;
                    }                                             
                    else if (sub.Contains("Лід із фейсбук"))
                    {
                        if (body.Contains("Номер телефону: "))
                            ph = TakePhone(body, "Номер телефону: ");

                        if (body.Contains("Кава як"))
                            whatType = TakeLineFromMail(body, "Кава як", 50);
                        else
                            whatType = "";

                        if (body.Contains("Ім’я: "))
                            pib = TakeName(body, "Ім’я: ");
                        else
                            pib = "";

                        if (body.Contains("Область: "))
                            region = TakeLineFromMail(body, "Область: ", 250);
                        else
                            region = "";

                        if (body.Contains("Суть звернення: "))
                            Comment = TakeLineFromMail(body, "Суть звернення: ", 500);
                        else
                            Comment = "";

                        typeOfComplaint = "Надання комерційної пропозиції";

                        mailType = 5;
                    }
                    else
                        mailType = 0;

                    log.WriteLine("Saving letter info to base.");


                    if (mailType > 0 && mailType < 5)
                    {
                        List<string> typeColumns = new List<string> { };
                        List<string> typeParams = new List<string> { };
                        if (mailType == 1)
                        {
                            typeColumns = new List<string> { "Subject", "email", "Body", "Gender", "FullName", "phone", "CompanyName", "PostCode", "Comment", "MachineCode", "TransfPrefixTail", "BodyHTML" };
                            typeParams = new List<string> { sub, EmailAddr, body, Gender, pib, ph, CompName, Postcode, Comment, MachineCode, "3", bodyHTML };
                        }
                        else if (mailType == 2)
                        {
                            typeColumns = new List<string> { "Subject", "email", "Body", "Gender", "FullName", "CompanyName", "MachineCode", "phone", "Address", "Question", "Comment", "Agreement", "TransfPrefixTail", "BodyHTML" };
                            typeParams = new List<string> { sub, EmailAddr, body, Gender, pib, CompName, MachineCode, ph, address, dw, Comment, Ag.ToString(), "3", bodyHTML };
                        }

                        if (typeColumns.Count > 0)
                            CmdExecute("zd_JacobsProfessional_RegMail", typeColumns, typeParams, false);
                    }
                    else if (mailType == 5)
                    {
                        List<string> type5Columns = new List<string> { "WhatType", "Phone", "FName", "txtReg", "Comment", "TypeOfComplaint", "LeadFrom" };
                        List<string> type5Params = new List<string> { whatType, ph, pib, region, Comment, typeOfComplaint, "соц. мережі" };
                        int infoID = CmdExecute("zd_JacobsProfessional_Info", type5Columns, type5Params, true);
                        //infoID = cmdExecute("zd_JacobsProfessional_Info", new string[] { "WhatType", "Phone", "FName", "txtReg", "Comment", "TypeOfComplaint" }, new string[] { whatType, ph, pib, region, Comment, typeOfComplaint });

                        type5Columns.Add("INFOID");
                        type5Params.Add(infoID.ToString());
                        CmdExecute("zd_JacobsProfessional_Reg", type5Columns, type5Params, false);

                        var client = new WebClient();
                        var content = client.DownloadString(ConfigurationManager.AppSettings["statLink"] + infoID.ToString());
                        log.WriteLine("Send mail link is executed. Result: " + content);
                    }

                    message.IsRead = true;
                    message.Update(ConflictResolutionMode.AlwaysOverwrite);
                    message.Move(folderProcessedId);
                    log.WriteLine("Took a letter!");
                }
            }
        }

        //Forms and executes SQL query in format 'INSERT INTO [TABLE](A,B,C) VALUES(A1, B1, C1)'
        //Takes params queryTable=TABLE, tableColums=(A,B,C), passingParameters=(A1,B1,C1)
        private static int CmdExecute(string queryTable, List<string> tableColumns, List<string> passingParameters, bool getOutput)
        {
            int returnedID = 0;
            using (SqlConnection sqlCon = new SqlConnection(ConfigurationManager.ConnectionStrings["myCon"].ConnectionString))
            {
                string fullCommand = "set dateformat dmy insert into [dbo].[" + queryTable + "](";
                for (int i = 0; i < tableColumns.Count; i++)
                    fullCommand += "[" + tableColumns[i] + "],";
                if (getOutput)
                    fullCommand = fullCommand.Remove(fullCommand.Length - 1, 1).Insert(fullCommand.Length - 1, ") output INSERTED.ID values(");//statement "output" cannot be used with tables that have triggers
                else
                    fullCommand = fullCommand.Remove(fullCommand.Length - 1, 1).Insert(fullCommand.Length - 1, ") values(");
                for (int i = 0; i < passingParameters.Count; i++)
                    fullCommand += "'" + passingParameters[i] + "',";
                fullCommand = fullCommand.Remove(fullCommand.Length - 1, 1).Insert(fullCommand.Length - 1, ")");

                SqlCommand cmd = new SqlCommand();
                cmd = new SqlCommand(fullCommand, sqlCon);

                if (sqlCon.State != ConnectionState.Open)
                    sqlCon.Open();
                if (getOutput)
                    returnedID = (int)cmd.ExecuteScalar();
                else
                    cmd.ExecuteNonQuery();
            }

            return returnedID;
        }

        private static FolderId GetExchangeFolderIdByName(string folderName, ExchangeService service, Mailbox mbox)
        {
            FolderId fid = new FolderId(WellKnownFolderName.Inbox, mbox);
            FolderView view = new FolderView(100)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly) { FolderSchema.DisplayName },
                Traversal = FolderTraversal.Deep
            };
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.Root, view);
            foreach (Folder f in findFolderResults)
            {
                if (f.DisplayName == folderName)
                    fid = f.Id;
            }
            return fid;
        }

        private static string TakeLineFromMail(string MailBody, string StartPoint, int LineLength)
        {
            try
            {
                string processedText;
                processedText = MailBody.Substring(MailBody.IndexOf(StartPoint));
                if (processedText.IndexOf("\n") > 0)
                    processedText = processedText.Remove(processedText.IndexOf("\n"));
                if (processedText.IndexOf("\r") > 0)
                    processedText = processedText.Remove(processedText.IndexOf("\r"));
                if (processedText.Contains(":"))
                    processedText = processedText.Remove(0, processedText.IndexOf(":") + 1);
                if (processedText.Length > LineLength)
                    processedText = processedText.Substring(0, LineLength);
                processedText.Trim();

                return processedText;
            }
            catch
            {
                return "";
            }
        }

        private static string TakePhone(string MailBody, string phoneTag)
        {
            string ph = TakeLineFromMail(MailBody, phoneTag, 20); //function takes from the left side, whereas phone needs to be taken from the right side, so i take extra symbols, and then cut leftovers
            ph = (Regex.Replace(ph, @"[^\d]", "", RegexOptions.Compiled)).Right(10);

            return ph;
        }

        private static string TakeName(string MailBody, string nameTag)
        {
            string pib = TakeLineFromMail(MailBody, nameTag, 250);
            pib = Regex.Replace(pib, @"^[a-zA-Z]+$", "");

            return pib;
        }
    }
}
