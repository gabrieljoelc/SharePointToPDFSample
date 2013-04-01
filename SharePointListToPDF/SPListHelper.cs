using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Protocols;
using System.Xml;
using System.Xml.Linq;

namespace SharePointListToPDF
{
    public static class SPListHelper
    {
        public static IEnumerable<XElement> GetListItems(Uri url = null, string listName = null, string viewName = null)
        {
            XNamespace z = "#RowsetSchema";
            url = url ?? SharepointConfig.WebServiceUrl;
            listName = listName ?? SharepointConfig.ListName;
            viewName = viewName ?? SharepointConfig.ViewName;

            XmlNode textSpsList;
            using (var ws = new SPListProxy(url))
            {
                textSpsList = ws.GetListItems(listName, viewName);
            }

            XElement elements = GetXElement(textSpsList);
            return elements.Descendants(z + "row");
        }

        private static XElement GetXElement(XmlNode node)
        {
            var xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
            {
                node.WriteTo(xmlWriter);
            }
            return xDoc.Root;
        }

        [WebServiceBinding(Name = "ListsSoap", Namespace = "http://schemas.microsoft.com/sharepoint/soap/")]
        private class SPListProxy : SoapHttpClientProtocol
        {
            public SPListProxy(Uri uri)
            {
                Url = uri.ToString();
                Credentials = SharepointConfig.Credentials;
            }

            [SoapDocumentMethod("http://schemas.microsoft.com/sharepoint/soap/GetListItems",
                RequestNamespace = "http://schemas.microsoft.com/sharepoint/soap/",
                ResponseNamespace = "http://schemas.microsoft.com/sharepoint/soap/", Use = SoapBindingUse.Literal,
                ParameterStyle = SoapParameterStyle.Wrapped)]
            public XmlNode GetListItems(string listName, string viewName = null, XmlNode query = null,
                XmlNode viewFields = null,
                string rowLimit = null, XmlNode queryOptions = null, string webID = null)
            {
                object[] results = Invoke("GetListItems", new object[]
                {
                    listName,
                    viewName,
                    query,
                    viewFields,
                    rowLimit,
                    queryOptions,
                    webID
                });
                return ((XmlNode)(results[0]));
            }
        }

    }

    static class SharepointConfig
    {
        public static ICredentials Credentials
        {
            get
            {
                var username = ConfigurationManager.AppSettings.Get("sp_username");
                var password = ConfigurationManager.AppSettings.Get("sp_password");
                var domain = ConfigurationManager.AppSettings.Get("sp_domain");
                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    return new NetworkCredential(username, password, domain);
                }
                return CredentialCache.DefaultCredentials;
            }
        }

        public static Uri WebServiceUrl
        {
            get
            {
                return new Uri(ConfigurationManager.AppSettings.Get("sp_endpoint"));
            }
        }

        public static string ListName
        {
            get
            {
                return ConfigurationManager.AppSettings.Get("sp_listName");
            }
        }

        public static string ViewName
        {
            get
            {
                return ConfigurationManager.AppSettings.Get("sp_viewName");
            }
        }
    }
}