using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using SharePoint.ReciveMailAddinWeb.Models;
using Microsoft.SharePoint.Client;
using System.Security;
using System.IO;
using System.Configuration;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using System.Text;
using System.Net.Mail;

namespace SharePoint.ReciveMailAddinWeb.Controllers
{
    public class MailController : ApiController
    {
        [HttpPost]
        public async Task<HttpResponseMessage> Post()
        {
            var root = System.IO.Path.GetTempPath();
            var provider = new MultipartFormDataStreamProvider(root);
            try
            {
                await Request.Content.ReadAsMultipartAsync(provider);
                SaveMail(provider);
            }
            catch (Exception ex)
            {
                var error = Path.Combine(HttpContext.Current.Server.MapPath("~/App_Data"),"ERR.TXT");
                System.IO.File.AppendAllText(error, "\n\n" + DateTime.Now.ToString() +"\n" + ex.ToString());
                foreach(var item in provider.FormData.AllKeys)
                {
                    System.IO.File.AppendAllText(error, "\n" + item + "=" + (GetProviderValue(provider, item) ??"null") );
                }
//                throw;
            }
            return new HttpResponseMessage(HttpStatusCode.OK);
        }


        private void SaveMail(MultipartFormDataStreamProvider provider)
        {
            var files = provider.FileData;
            var email = new Email
            {
                Dkim = GetProviderValue(provider,"dkim"),
                To = GetProviderValue(provider, "to"),
                Html = GetProviderValue(provider, "html"),
                From = GetProviderValue(provider, "from"),
                Text = GetProviderValue(provider, "text"),
                SenderIp = GetProviderValue(provider, "sender_ip"),
                Envelope = GetProviderValue(provider, "envelope"),
                Subject = GetProviderValue(provider, "subject"),
                Charsets = GetProviderValue(provider, "charsets"),
                Spf = GetProviderValue(provider, "spf")
            };

            SaveMail(files, email);
        }

        private string GetProviderValue(MultipartFormDataStreamProvider provider, string key)
        {
            var val = provider.FormData.GetValues(key);
            return val == null ? null : val.FirstOrDefault();
        }

        private void SaveMail(Collection<MultipartFileData> files, Email email)
        {
            var siteUrl = ConfigurationManager.AppSettings.Get("ReciveMail.SiteUrl");
            var user = ConfigurationManager.AppSettings.Get("ReciveMail.AccessUser");
            var pass = ConfigurationManager.AppSettings.Get("ReciveMail.AccessPassword");

            var settings = (MailSettings)HttpContext.Current.Application["MailSettings"];
            
            using (var ctx = CreateContext(siteUrl, user, pass))
            {
                var envelope = JsonConvert.DeserializeObject<dynamic>(email.Envelope);
                foreach (string toUser in envelope.to)
                {
                    if (settings == null || settings.Rules == null) continue;

                    var toUserAddress = new MailAddress(toUser);

                    var rule = settings.Rules.Where(x => x.MailAddress.Equals(toUserAddress.Address, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (rule != null)
                    {
                        var list = ctx.Web.Lists.GetByTitle(rule.ListTile);

                        if (list == null) continue;
                        var charsetsObj = JsonConvert.DeserializeObject<dynamic>(email.Charsets);

                        var itemCreateInfo = new ListItemCreationInformation();
                        var item = list.AddItem(itemCreateInfo);
                        item["Title"] = ConvertToUtf8(email.Subject, Encoding.GetEncoding((string)charsetsObj.subject)); 

                        if (!string.IsNullOrEmpty(email.Html))
                        {
                            var encode = Encoding.GetEncoding((string)charsetsObj.html);
                            var decodedHtml = ConvertToUtf8(email.Html, encode);
                            item["Body"] = $"From:{ConvertToUtf8(email.From, Encoding.GetEncoding((string)charsetsObj.from))}<hr/>{decodedHtml}";
                        }
                        else
                        {
                            var encode = Encoding.GetEncoding((string)charsetsObj.text);
                            string decodedText = ConvertToUtf8(email.Text, encode);
                            item["Body"] = $"From:{ConvertToUtf8(email.From, Encoding.GetEncoding((string)charsetsObj.from))}<hr/>{ decodedText.Replace("\n", "<br />") }";
                        }
                        item.Update();
                        ctx.ExecuteQuery();

                        if (files != null && files.Count > 0)
                        {
                            foreach (var fileData in files)
                            {
                                if (string.IsNullOrEmpty(fileData.Headers.ContentDisposition.FileName))
                                {
                                    throw new Exception("This request is not properly formatted");
                                }
                                if (new FileInfo(fileData.LocalFileName).Length == 0) continue;
                                string fileName = fileData.Headers.ContentDisposition.FileName;
                                if (fileName.StartsWith("\"") && fileName.EndsWith("\""))
                                {
                                    fileName = fileName.Trim('"');
                                }
                                if (fileName.Contains(@"/") || fileName.Contains(@"\"))
                                {
                                    fileName = Path.GetFileName(fileName);
                                }
                                fileName = DecodeSendGridAttachementFile(fileName);
                                var attCreateInfo = new AttachmentCreationInformation();
                                attCreateInfo.FileName = fileName;
                                using (var st = new FileStream(fileData.LocalFileName, FileMode.Open))
                                {
                                    attCreateInfo.ContentStream = st;
                                    var atta = item.AttachmentFiles.Add(attCreateInfo);
                                    ctx.Load(atta);
                                    ctx.ExecuteQuery();
                                }
                            }
                        }
                    }
                }
            }
        }

        private string ConvertToUtf8(string text, Encoding encode)
        {
            var original = encode.GetBytes(text);
            var utf8 = Encoding.Convert(encode, Encoding.UTF8, original);
            var decodedText = Encoding.UTF8.GetString(utf8);
            return decodedText;
        }

        private string DecodeSendGridAttachementFile(string name)
        {
            var ret = new StringBuilder();
            for (var idx = 0; idx < name.Length; idx++)
            {
                if (name[idx] == '%' && idx + 4 < name.Length)
                {
                    var unicode = Convert.ToInt32(name.Substring(idx + 1, 4), 16);
                    char ch = Convert.ToChar(unicode);
                    ret.Append(ch);
                    idx += 4;
                }
                else
                {
                    ret.Append(name[idx]);
                }
            }
            return ret.ToString();
        }


        private static ClientContext CreateContext(string siteUrl, string user, string password)
        {
            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            var onlineCredentials = new SharePointOnlineCredentials(user, securePassword);
            var context = new ClientContext(siteUrl);
            context.Credentials = onlineCredentials;
            return context;

        }
    }
}