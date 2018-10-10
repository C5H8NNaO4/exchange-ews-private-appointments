using System;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using NDesk.Options;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace privApt
{
    static class Constants
    {
        public const bool DEBUG_DONT_UPDATE = true;
        public const int DEBUG_PAGES = 100000;

    }
    class Program
    {
        static int pageSize = 5;
        static int processed = 0;

        static ExchangeService service;

        static String discoverUser = "";
        static String serviceUser = "";
        static String servicePass = "";

        static string[] users = new string[0]; //"elmar.schoettner@bechtle.com,moritz.roessler@bechtle.com".Split(',');
        static int verbosity = 0;
        static string path = "";
        static string url = "";
        static bool save;
        static bool normal;
        static int tim = 15;
        static string logfile = System.IO.Path.GetTempPath() + "privApts.log";
        static StreamWriter w;
        static bool show_help = false;
        static bool parallelUsers;
        static bool parallelPages;
        static int total = 0;
        static void Main(string[] args)
        {
            OptionSet p = initOptions(args);
            if (File.Exists(logfile)) File.Delete(logfile);
            w = File.AppendText(logfile);
            logOptions(args);
            initialize(p);



            Log("Autodiscover URL " + service.Url);
            if (parallelUsers) Parallel.ForEach(users, processCalendar);
            else
            {
                for (int i = 0; i < users.Length; i++)
                {
                    processCalendar(users [i]);
                }
            }
            Log("Processed a total of " + Program.total + " appointments in " + users.Length + " calendars.");
            LogLine("Logfile at: " + logfile);

            exit();

        }
 
        static void processCalendar (string usr) {
            int total = 0;
            int page = 0;
            LogLine("Processing Calendar: " + usr);
            FindItemsResults<Item> results;
            do
            {
                results = loadPage(usr, page++);
                if (results == null)
                {
                    LogLine("No appointments found for user " + usr);
                    continue;
                }
                LogLine("Page " + page + "("+usr+"): " + results.Items.Count + " appointments.");
                turnPrivate(results);
                total += results.Items.Count;

            } while (Constants.DEBUG_PAGES > page && results != null && results.MoreAvailable);
            Program.total += total;
            LogLine(usr + " - processed " + total + " appointments.");
        }
        static void logOptions (string[] args)
        {
            LogLine("private_appointments " + String.Join(" ", args));
            LogStart();
            LogLine("users=" + path.ToString());
            LogLine("save=" + save);
            LogLine("logfile=" + logfile);
            LogLine("discover=" + discoverUser.ToString());
            LogLine("user=" + serviceUser.ToString());
            LogLine("pass=" + "*****");
            LogLine("verbosity=" + verbosity);
            LogLine("normal=" + normal);
        }
        static OptionSet initOptions(string[] args)
        {
            var p = new OptionSet() {
                { "i|users=", "the {PATH} to the file, containing SMPT addresses.",
                   v => path = v },
                { "l|log=",
                   "the {PATH} to the file used for logging.",
                    (v) => logfile = v },
                { "c|page-size=",
                   "Number of items per page",
                    (v) => pageSize = int.Parse(v) },
                { "X+|parallel-pages",
                   "Number of items per page",
                    (v) => parallelPages = v!= null },
                { "U+|parallel-users",
                   "Number of items per page",
                    (v) => parallelUsers = v!= null },
                { "s+|save",
                   "Save the changes",
                    (v) => save = v != null},
                { "d|discover=",
                   "the {EMAIL} used to find the autodiscover url.",
                    (v) => discoverUser = v },
                { "u|user=",
                   "the {EMAIL} of the service user.",
                    (v) => serviceUser = v },
                { "n+|normal",
                   "Set appointments to normal.",
                    (v) => normal = v !=null },
                { "p|pass=",
                   "the {PASSWORD} for the service user",
                    (v) => servicePass = v },
                { "v", "Increase the verbosity",
                   v => { if (v != null) ++verbosity; } },
                { "e|url=", "the EWS endpoint {URL}.",
                   v => { url = v; }},
                { "h|help",  "show this message and exit",
                   v => show_help = v != null },
                { "timeout=",
                   "The timeout between saves.",
                    (v) => tim = int.Parse(v) },
            };

            try
            {
                p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("private_appointments: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try privApppointments '--help' for more information.");
                exit();
            }

            return p;
        }
        static void initialize (OptionSet p)
        {
            if (path == "")
            {
                LogLine("Error: Missing required option 'users'");
                show_help = true;
            }

            else
            {
                users = File.Exists(path) ? (File.ReadAllLines(path)) : new string[0];
                service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                service.Credentials = new WebCredentials(serviceUser, servicePass);
                if (verbosity > 3)
                    service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;
            }
            if (serviceUser == "")
            {
                LogLine("Error: Missing required option 'user'");
                show_help = true;
            }
            if (servicePass == "")
            {
                LogLine("Error: Missing required option 'pass'");
                show_help = true;
            }
            if (url == "")
            {
                if (discoverUser == "")
                {
                    LogLine("Error: Missing required option 'discover'");
                    w.Flush();
                    show_help = true;

                }
                else
                {
                    try
                    {
                        service.AutodiscoverUrl(discoverUser, RedirectionUrlValidationCallback);
                    }
                    catch (Exception e)
                    {
                        LogLine(e.Message);
                        Log("Autodiscover Failed");
                        w.Flush();
                        return;
                    }
                }

            }
            else
            {
                LogLine("AutoDiscover Disabled, using " + url + " as endpoint");
                service.Url = new Uri(url);
            }

            if (show_help)
            {
                ShowHelp(p);
                exit();
                return;
            }
        }
        static void exit ()
        {
            w.Flush();
            w.Close();
            using (StreamReader r = File.OpenText(logfile))
            {
                DumpLog(r);
                Environment.Exit(0);
            }
        }
        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage: private_appointments [OPTIONS]+");
            Console.WriteLine("Set the sensitivity of every appointment to 'Private'.");
            Console.WriteLine("Expects a list of users.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        public static void LogStart()
        {
            w.Write("\r\nLog Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                DateTime.Now.ToLongDateString());
            w.WriteLine("  :");
        }
        public static void LogLine(string logMessage)
        {
            if (verbosity > 1) Console.WriteLine(logMessage);
            w.WriteLine("  :{0}", logMessage);

        }
        public static void LogEnd()
        {
            w.WriteLine("-------------------------------");
        }
        public static void Log(string logMessage)
        {
            LogStart();
            w.WriteLine("  :{0}", logMessage);
            LogEnd();
        }

        public static void DumpLog(StreamReader r)
        {
            string line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }
        private static void turnPrivate(FindItemsResults<Item> results)
        {
            if (parallelPages) Parallel.ForEach(results.Items, updateItem);
            else
            {
                for (int i=0; i < results.Items.Count; i++)
                {
                    updateItem(results.Items[i]);
                }
            }

        }

        private static void updateItem (Item apt)
        {
            Appointment item = (Appointment)apt;

            if (verbosity > 1)
            {
                 Console.WriteLine(++processed + " " + item.Subject);
            }
                    
            item.Sensitivity = normal ? Sensitivity.Normal : Sensitivity.Private;
            try
            {
                if (save)
                {
                    if (verbosity > 0) LogLine("Updating " + item.Sensitivity+ " item '" + item.Subject + "'");
                    item.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendToNone);
                }

                if (verbosity > 2) LogLine("    Id " + item.Id.ToString());
                System.Threading.Thread.Sleep(tim);
            }
            catch (Exception e)
            {
                Log("Error Updating Item : " + item.Subject);
                LogLine("Error Updating Item : " + e.Message);

            }
        }
        private static FindItemsResults<Item> loadPage(string usr, int pageNr)
        {
            int offset = (pageSize * pageNr);

            ItemView view = new ItemView(pageSize , offset);

            return fetchAppointments(usr, view);
        }

        private static FindItemsResults<Item> fetchAppointments(string usr, ItemView view)
        {

            view.PropertySet = new PropertySet(ItemSchema.Subject,
                                   ItemSchema.DateTimeReceived,
                                   EmailMessageSchema.IsRead);
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
            view.Traversal = ItemTraversal.Shallow;
            try
            {
                // Call FindItems to find matching calendar items. 
                // The FindItems parameters must denote the mailbox owner,
                // mailbox, and Calendar folder.
                // This method call results in a FindItem call to EWS.
                FindItemsResults<Item> results = service.FindItems(
                new FolderId(WellKnownFolderName.Calendar,
                    usr),
                    view);
                return results;

            }
            catch (Exception ex)
            {
                LogLine("Error fetching appointments for user: " + usr);
                LogLine(ex.Message);
            }
            return null;

        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}