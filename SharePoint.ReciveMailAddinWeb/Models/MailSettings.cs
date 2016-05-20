using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace SharePoint.ReciveMailAddinWeb.Models
{
    public class MailItem
    {
        public string MailAddress { get; set; }

        public string ListTile { get; set; }

    }
    public class MailSettings
    {
        public List<MailItem> Rules { get; set; }


        public MailSettings()
        {
            Rules = new List<MailItem>();        
        }

        public static MailSettings ReadSettings(string data)
        {
            MailSettings settings = new MailSettings();
            if (System.IO.File.Exists(data))
            {
                var ser = new XmlSerializer(typeof(MailSettings));
                using (var st = new StreamReader(data))
                {
                    settings = (MailSettings)ser.Deserialize(st);

                }
            }

            return settings;
        }


        public void Save(string data)
        {
            var ser = new XmlSerializer(typeof(MailSettings));
            using (var st = new StreamWriter(data))
            {
                ser.Serialize(st, this);
            }
            HttpContext.Current.Application["MailSettings"] = data;
        }

    }
}