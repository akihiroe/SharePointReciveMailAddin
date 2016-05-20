using Microsoft.SharePoint.Client;
using SharePoint.ReciveMailAddinWeb.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Web;
using System.Web.Mvc;
using System.Xml.Serialization;

namespace SharePoint.ReciveMailAddinWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            var data = Server.MapPath("~/App_Data/MailSettings.xml");
            MailSettings vm = MailSettings.ReadSettings(data);
            vm.Rules.Add(new MailItem());
            return View(vm);
        }

        public ActionResult Update(MailSettings vm)
        {
            var data = Server.MapPath("~/App_Data/MailSettings.xml");
            for (int i=vm.Rules.Count()-1; i>=0; i--)
            {
                if (string.IsNullOrWhiteSpace(vm.Rules[i].MailAddress) || string.IsNullOrWhiteSpace(vm.Rules[i].ListTile)) vm.Rules.RemoveAt(i);
            }
            vm.Save(data);
            vm.Rules.Add(new MailItem());
            return View("Index", vm);
        }

    }
}
