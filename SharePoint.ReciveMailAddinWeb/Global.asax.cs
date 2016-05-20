using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Http;
using SharePoint.ReciveMailAddinWeb.Models;

namespace SharePoint.ReciveMailAddinWeb
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            var data = Server.MapPath("~/App_Data/MailSettings.xml");
            var settings = MailSettings.ReadSettings(data);
            this.Application["MailSettings"] = settings;
        }
    }
}
