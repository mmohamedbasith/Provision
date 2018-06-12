using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Net;
using System.Security;
using System.Threading;

namespace Provision
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ConsoleColor defaultForeground = Console.ForegroundColor;

                // Collect information 
                // string templateWebUrl = GetInput("Enter the URL of the template site: ", false, defaultForeground);
                string targetWebUrl = GetInput("Enter the URL of the target site: ", false, defaultForeground);
                string userName = GetInput("Enter your user name:", false, defaultForeground);
                string pwdS = GetInput("Enter your password:", true, defaultForeground);
                string filepath = GetInput("Get XMl Path:", false, defaultForeground);
                string filename = GetInput("Get XMl filename:", false, defaultForeground);

                SecureString pwd = new SecureString();

                foreach (char c in pwdS.ToCharArray()) pwd.AppendChar(c);
                using (var context = new ClientContext(targetWebUrl))
                {


                    context.Credentials = new SharePointOnlineCredentials(userName, pwd);
                    Web web = context.Web;
                    context.Load(web, w => w.Title);
                    context.ExecuteQueryRetry();
                    XMLTemplateProvider provider =
                             new XMLFileSystemTemplateProvider(filepath, "");


                    ProvisioningTemplate template = provider.GetTemplate(filename);
                    ProvisioningTemplateApplyingInformation ptai =
                    new ProvisioningTemplateApplyingInformation();
                    ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                    };



                    template.Connector = provider.Connector;

                    web.ApplyProvisioningTemplate(template, ptai);
                }
                // Get the template from existing site and serialize that (not really needed)
                // Just to pause and indicate that it's all done
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("We are all done. Press enter to continue.");
                Console.ReadLine();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static ProvisioningTemplate GetProvisioningTemplate(ConsoleColor defaultForeground, string webUrl, string userName, SecureString pwd)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is:" + ctx.Web.Title);
                Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector, so that we can store composed files temporarely somewhere 
                ptci.FileConnector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
                ptci.PersistBrandingFiles = true;
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Execute actual extraction of the tepmplate
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

                // We can also serialize this template for future usage if we want, not really needed
                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(@"c:\temp\pnpprovisioningdemo", "");
                provider.SaveAs(template, "PnPProvisioningDemo.xml");

                return template;
            }
        }

        private static void ApplyProvisioningTemplate(ConsoleColor defaultForeground, string webUrl, string userName, SecureString pwd, ProvisioningTemplate template)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is:" + ctx.Web.Title);
                Console.ForegroundColor = defaultForeground;

                // We could potentially also upload the template from file system, but we at least need this for branding file
                //XMLTemplateProvider provider =
                //       new XMLFileSystemTemplateProvider(@"c:\temp\pnpprovisioningdemo", "");
                //template = provider.GetTemplate("PnPProvisioningDemo.xml");

                ProvisioningTemplateApplyingInformation ptai
                        = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Associate file connector for assets
                FileSystemConnector connector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
                template.Connector = connector;

           

                web.ApplyProvisioningTemplate(template, ptai);
            }
        }

        private static string GetInput(string label, bool isPassword, ConsoleColor defaultForeground)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("{0} : ", label);
            Console.ForegroundColor = defaultForeground;

            string value = "";

            for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (value.Length > 0)
                    {
                        value = value.Remove(value.Length - 1);
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    }
                }
                else if (keyInfo.Key != ConsoleKey.Enter)
                {
                    if (isPassword)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(keyInfo.KeyChar);
                    }
                    value += keyInfo.KeyChar;

                }

            }
            Console.WriteLine("");

            return value;
        }
    }
}
