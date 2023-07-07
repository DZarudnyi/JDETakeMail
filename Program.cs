using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using System.Net;
using System.Collections.Generic;
using System.Configuration;

namespace JDETakeMail
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger log = new Logger(String.Format("JDETakeMail_log_{0}", DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss")), ConfigurationManager.AppSettings["logPath"]);// @"D:\logs\JDETakeMail"
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
            log.WriteLine("Configured mailbox.");
            FolderId folderProcessedId = GetExchangeFolderIdByName("Processed", service, mailBox, log);

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

                    ReplaceQuotes(ref addressTo);
                    ReplaceQuotes(ref sub);
                    ReplaceQuotes(ref body);
                    ReplaceQuotes(ref bodyHTML);

                    // Parsing taken letter
                    if (sub.Contains("Quote request online leads"))
                    {
                        pib = TakeName(body, "Name");
                        EmailAddr = TakeLineFromMail(body, "Email", 150);

                        ph = TakePhone(body, "PhoneNumber");
                        if (!ph.IsNumber())
                        {
                            DeleteMessage(message, log);
                            continue;
                        }

                        CompName = TakeLineFromMail(body, "CompanyName", 100);
                        Gender = TakeLineFromMail(body, "Gender", 20);
                        Postcode = Regex.Replace(TakeLineFromMail(body, "Postcode", 50), @"[^\d]", "", RegexOptions.Compiled);
                        Comment = TakeLineFromMail(body, "AdditionalComments", 500);
                        MachineCode = TakeLineFromMail(body, "MachineCode", 50);

                        mailType = 1;
                    }
                    else if (sub.Contains("Contact"))
                    {
                        pib = TakeName(body, "Name");
                        CompName = TakeLineFromMail(body, "Company name", 100);
                        EmailAddr = TakeLineFromMail(body, "E-mail", 150);

                        ph = TakePhone(body, "Phone Number");
                        if (!ph.IsNumber())
                        {
                            DeleteMessage(message, log);
                            continue;
                        }

                        Gender = TakeLineFromMail(body, "Salutation", 20);
                        MachineCode = TakeLineFromMail(body, "Coffee machine", 50);
                        address = TakeLineFromMail(body, "Adress", 250);//it is intended; it is mistake in client`s template
                        dw = TakeLineFromMail(body, "Question", 100);
                        Comment = TakeLineFromMail(body, "Your question", 500);
                        Ag = TakeLineFromMail(body, "Agreement", 50);

                        mailType = 2;
                    }                                             
                    else if (sub.Contains("Лід із фейсбук"))
                    {
                        ph = TakePhone(body, "Номер телефону");
                        if (!ph.IsNumber())
                        {
                            DeleteMessage(message, log);
                            continue;
                        }

                        whatType = TakeLineFromMail(body, "Кава як", 50);
                        pib = TakeName(body, "Ім’я");
                        region = TakeLineFromMail(body, "Область", 250);
                        Comment = TakeLineFromMail(body, "Суть звернення", 500);
                        typeOfComplaint = "Надання комерційної пропозиції";

                        mailType = 5;
                    }
                    else
                    {
                        mailType = 0;
                        DeleteMessage(message, log);
                        continue;
                    }
                        

                    if (ph.Substring(0, 3) == "555")
                    {
                        DeleteMessage(message, log);
                        continue;
                    }

                    log.WriteLine("Saving letter info to base.");

                    //meet field limitations before writing to db
                    if (body.Length > 4000)
                        body = body.Substring(0, 4000);
                    if (bodyHTML.Length > 4000)
                        bodyHTML = bodyHTML.Substring(0, 4000);


                    if (mailType > 0 && mailType < 5)
                    {
                        List<string> typeColumns = new List<string> { };
                        List<string> typeParams = new List<string> { };
                        if (mailType == 1)
                        {
                            typeColumns = new List<string> { "Subject", "email", "datetimeSent", "Body", "Gender", "FullName", "phone", "CompanyName", "PostCode", "Comment", "MachineCode", "TransfPrefixTail", "BodyHTML" };
                            typeParams = new List<string> { sub, EmailAddr, datetimeSent, body, Gender, pib, ph, CompName, Postcode, Comment, MachineCode, "3", bodyHTML };
                        }
                        else if (mailType == 2)
                        {
                            typeColumns = new List<string> { "Subject", "email", "datetimeSent", "Body", "Gender", "FullName", "CompanyName", "MachineCode", "phone", "Address", "Question", "Comment", "Agreement", "TransfPrefixTail", "BodyHTML" };
                            typeParams = new List<string> { sub, EmailAddr, datetimeSent, body, Gender, pib, CompName, MachineCode, ph, address, dw, Comment, Ag.ToString(), "3", bodyHTML };
                        }
                        log.WriteLine("Prepared list of values to write into table.");

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
                    log.WriteLine("Message proccessed!");
                }
            }
        }

        //Forms and executes SQL query in format 'INSERT INTO [TABLE](A,B,C) VALUES(A1, B1, C1)'
        //Takes params queryTable=TABLE, tableColums=(A,B,C), passingParameters=(A1,B1,C1)
        private static int CmdExecute(string queryTable, List<string> tableColumns, List<string> passingParameters, bool getOutput)
        {
            int returnedID = 0;

            ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings["con"];
            if (settings != null)
            {
                using (SqlConnection sqlCon = new SqlConnection(settings.ConnectionString))
                {
                    string fullCommand = "set dateformat dmy insert into [dbo].[" + queryTable + "](";
                    for (int i = 0; i < tableColumns.Count; i++)
                        fullCommand += "[" + tableColumns[i] + "],";
                    if (getOutput)
                        fullCommand = fullCommand.Remove(fullCommand.Length - 1, 1).Insert(fullCommand.Length - 1, ") output INSERTED.ID values(");//statement "output" cannot be used with tables that have triggers
                    else
                        fullCommand = fullCommand.Remove(fullCommand.Length - 1, 1).Insert(fullCommand.Length - 1, ") values(");
                    for (int i = 0; i < passingParameters.Count; i++)
                    {
                        if (passingParameters[i].Contains("'"))
                            passingParameters[i] = passingParameters[i].Replace("'", "`");
                        fullCommand += "'" + passingParameters[i] + "',";
                    }
                    fullCommand = fullCommand.Remove(fullCommand.Length - 1, 1).Insert(fullCommand.Length - 1, ")");

                    using (SqlCommand cmd = new SqlCommand(fullCommand, sqlCon))
                    {
                        //cmd.Parameters.Add("@1", SqlDbType.NVarChar, length).Value = someValue;
                        if (sqlCon.State != ConnectionState.Open)
                            sqlCon.Open();
                        if (getOutput)
                            returnedID = (int)cmd.ExecuteScalar();
                        else
                            cmd.ExecuteNonQuery();
                    }
                }
            }

            return returnedID;
        }

        private static FolderId GetExchangeFolderIdByName(string folderName, ExchangeService service, Mailbox mbox, Logger log)
        {
            FolderId fid = new FolderId(WellKnownFolderName.Inbox, mbox);
            FolderView view = new FolderView(100)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly) { FolderSchema.DisplayName },
                Traversal = FolderTraversal.Deep
            };
            //server seems to have problem with next line; probably it is because of service object
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.Root, view);
            foreach (Folder f in findFolderResults)
            {
                if (f.DisplayName == folderName)
                {
                    fid = f.Id;
                    break;
                }
            }
            return fid;
        }

        private static string ReplaceQuotes(ref string text)
        {
            if (text == null)
                return "";
            else
                return text.Replace("'", "''");
        }

        private static string TakeLineFromMail(string mailBody, string key, int lineLength)
        {
            if (mailBody.Contains(key))
            {
                try
                {
                    string processedText;
                    processedText = mailBody.Substring(mailBody.IndexOf(key));
                    if (processedText.IndexOf("\n") > 0)
                        processedText = processedText.Remove(processedText.IndexOf("\n"));
                    if (processedText.IndexOf("\r") > 0)
                        processedText = processedText.Remove(processedText.IndexOf("\r"));
                    if (processedText.Contains(":"))
                        processedText = processedText.Remove(0, processedText.IndexOf(":") + 1);
                    if (processedText.Length > lineLength)
                        processedText = processedText.Substring(0, lineLength);
                    processedText.Trim();

                    return processedText;
                }
                catch
                {
                    return "";
                }
            }
            else
                return "";
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

        private static void DeleteMessage(EmailMessage message, Logger log)
        {
            message.IsRead = true;
            message.Update(ConflictResolutionMode.AlwaysOverwrite);
            message.Move(WellKnownFolderName.DeletedItems);
            log.WriteLine("Found a spam letter. Deleted the letter.");
        }
    }
}
