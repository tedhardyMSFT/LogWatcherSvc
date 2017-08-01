using System;
using System.Collections.Generic;
using System.Configuration.Install;
using System.Diagnostics;
using System.ServiceProcess;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace LogWatcherSvc
{
    class LogWatcher 
    {
        /// <summary>
        /// Event Log legacy Event Source name for application events.
        /// </summary>
        internal static string ApplicationName = "LogWatcherSvc";

        /// <summary>
        /// Performance catgory name for Event Log related counters.
        /// </summary>
        internal static string PerformanceCounterCategoryName = "Windows Event Log";

        /// <summary>
        /// Name of the performace counter for Event Log
        /// </summary>
        internal static string PerformanceCounterName = "Channel EPS rate";
        
        /// <summary>
        /// Performance category name for Windows Event Forwarding/Collector counters.
        /// </summary>
        internal static string EventForwardingCounterCategoryName = "Windows Event Forwarding";

        /// <summary>
        /// Names of the Performance counters for Windows Event Forwarding category.
        /// </summary>
        internal static string[] EventForwardingCounterName = { 
                                                                  "Active Event Source Count", 
                                                                  "Total Event Source Count" 
                                                              };
        /// <summary>
        /// Descriptions for the 
        /// </summary>
        internal static string[] EventForwardingCounterDescriptions = {
                                                                          "Total count of hosts that are currently active in forwarding events to this subscription, these hosts have recently sent heartbeat messages or sent events.",
                                                                          "Total count of all hosts active and inactive that the subscription currently has heartbeat and bookmark records."
                                                                      };


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            if (System.Environment.UserInteractive)
            {
                if (args.Length > 0)
                {
                    switch (args[0].ToLower())
                    {
                        case "-i":
                            {   
                                // a crappy way to install a service, but I cannot find/figure out another way.
                                ManagedInstallerClass.InstallHelper(new string[] { Assembly.GetExecutingAssembly().Location });

                                // add the event log source name for all logging.
                                try
                                {
                                    if (!EventLog.SourceExists(LogWatcher.ApplicationName))
                                    {
                                        EventLog.CreateEventSource(LogWatcher.ApplicationName, "Application");
                                        Console.WriteLine("Created event source LogWatcherSvc in the Application log.");
                                    }
                                } // try
                                catch (InvalidOperationException AEx)
                                {
                                    // only expected error - source already exists. one of the scenarios for ArgumentException
                                    Console.WriteLine("Error creating event source: {1} in Application log. Message:{0}", 
                                        AEx.Message, 
                                        LogWatcher.ApplicationName
                                        );
                                    Console.WriteLine("Exiting.");

                                    Environment.Exit(1);
                                }

                                try
                                {
                                    if (!(PerformanceCounterCategory.Exists(LogWatcher.PerformanceCounterCategoryName)))
                                    {
                                        //// create the counter information, it will be used below in the Category creation
                                        //CounterCreationData EpsCounter = new CounterCreationData(
                                        //    LogWatcher.PerformanceCounterName,
                                        //    "Number of events written by a channel in a second",
                                        //    PerformanceCounterType.NumberOfItems32
                                        //    );

                                        //// perf counter category createion expects a collection, so give it one.
                                        //CounterCreationDataCollection CategoryCounters = new CounterCreationDataCollection();

                                        //// add the counter to the collection
                                        //CategoryCounters.Add(EpsCounter);

                                        //// Register Performance Coutner Category (aka Object name in PerfMon)
                                        //// the CategoryCounters is a collection of the Counters to be created (aka Counter in Perfmon)
                                        //// the application will only add/remove instances to the counter.
                                        //PerformanceCounterCategory.Create(LogWatcher.PerformanceCounterCategoryName,
                                        //    "The Windows Event Log performance object consists of counters that measure aspects of Event log activity. The Event log is the part of the computer that logs actions performed by the computer, whether on behalf of itself or a user ccount.  A computer has multiple event logs, each called a Channel.",
                                        //    PerformanceCounterCategoryType.MultiInstance,
                                        //    CategoryCounters
                                        //    );


                                        PerformanceCounterCategory.Create(
                                            LogWatcher.PerformanceCounterCategoryName,
                                            "The Windows Event Log performance object consists of counters that measure aspects of Event log activity. The Event log is the part of the computer that logs actions performed by the computer, whether on behalf of itself or a user account.  A computer has multiple event logs, each called a Channel.",
                                            PerformanceCounterCategoryType.MultiInstance,
                                            LogWatcher.PerformanceCounterName,
                                            "Number of events written by a channel in a second"
                                            );

                                        System.Diagnostics.EventLog.WriteEntry(
                                            LogWatcher.ApplicationName,
                                                string.Format(
                                                    "Successfully added counter name:{0}",
                                                    LogWatcher.PerformanceCounterName),
                                                EventLogEntryType.Information,
                                                10001
                                                );

                                        Console.WriteLine("Created Performance counter object: {0}", LogWatcher.PerformanceCounterCategoryName);
                                    } // if (!(System.Diagnostics.PerformanceCounterCategory.Exists("Windows Event Log")))

                                    // add the Windows Event Forwarding category
                                    if(!(PerformanceCounterCategory.Exists(LogWatcher.EventForwardingCounterCategoryName)))
                                    {
                                        CounterCreationDataCollection WefCounters = new CounterCreationDataCollection();
                                        // add each counter under it.
                                        for(int i = 0; i < LogWatcher.EventForwardingCounterName.Length; i++)
                                        {

                                            CounterCreationData WefCounter = new CounterCreationData(
                                                LogWatcher.EventForwardingCounterName[i],
                                                LogWatcher.EventForwardingCounterDescriptions[i],
                                                PerformanceCounterType.NumberOfItems32);

                                            // add it to the collection
                                            int counterIndex = WefCounters.Add(WefCounter);

                                            System.Diagnostics.EventLog.WriteEntry(
                                                LogWatcher.ApplicationName,
                                                    string.Format(
                                                        "Successfully added Performance Counter name:{0} with index position:{1}",
                                                        LogWatcher.EventForwardingCounterCategoryName,
                                                        counterIndex),
                                                    EventLogEntryType.Information,
                                                    10000
                                                    );
                                        }

                                        // Create the counter category using the collection of counter creation objects above.
                                        PerformanceCounterCategory.Create(
                                            LogWatcher.EventForwardingCounterCategoryName,
                                            "The Windows Event Forwarding performance object consists of counters that measure aspects of Windows Event Forwarding subscription activity. Windows Event Forwarding allows selected events to be sent (by either push or pull subscription) to another Windows server for centralized storage.",
                                            PerformanceCounterCategoryType.MultiInstance,
                                            WefCounters
                                            );

                                        System.Diagnostics.EventLog.WriteEntry(
                                            LogWatcher.ApplicationName,
                                                string.Format(
                                                    "Successfully added Performance Counter category name:{0}",
                                                    LogWatcher.EventForwardingCounterCategoryName),
                                                EventLogEntryType.Information,
                                                10001
                                                );


                                    } // if(!(PerformanceCounterCategory.Exists(LogWatcher.EventForwardingCounterCategoryName))
                                }
                                catch (UnauthorizedAccessException UAEx)
                                {
                                    Console.WriteLine("Error creating Performance Counters. Message:{0}", UAEx.Message);
                                    Console.WriteLine("Exiting.");
                                    Environment.Exit(1);
                                }
                                break;
                            } // case "-i":
                        case "-u":
                            {
                                // attempting to uninstall a service that doesn't exist causes the InstallHelper to throw an exception.

                                // get list of installed services
                                ServiceController[] CurrentServices = System.ServiceProcess.ServiceController.GetServices();

                                foreach (ServiceController InstalledService in CurrentServices)
                                {
                                    // find a matching service name, then uninstall it.
                                    if (InstalledService.ServiceName == LogWatcher.ApplicationName)
                                    {
                                        try
                                        {
                                            // a crappy way to UN-install a service, but I cannot find/figure out another way.
                                            ManagedInstallerClass.InstallHelper(new string[] { "/u", Assembly.GetExecutingAssembly().Location });
                                        }
                                        catch (System.Configuration.Install.InstallException)
                                        {
                                            // do something useful here.
                                        }
                                    } // if (InstalledService.ServiceName == LogWatcher.ApplicationName)
                                } // foreach (ServiceController InstalledService in CurrentServices)

                                try
                                {
                                    // remove event source from the server
                                    if (EventLog.SourceExists(LogWatcher.ApplicationName))
                                    {
                                        EventLog.DeleteEventSource(LogWatcher.ApplicationName);
                                    }
                                }
                                catch (System.Security.SecurityException )
                                {
                                    // need to do something useful here
                                }
                                catch (ArgumentException )
                                {
                                    // Ibid.
                                }

                                try
                                {
                                    // remove the Event Log performance counter
                                    if (PerformanceCounterCategory.Exists(LogWatcher.PerformanceCounterCategoryName))
                                    {
                                        PerformanceCounterCategory.Delete(LogWatcher.PerformanceCounterCategoryName);
                                    }

                                    // remove the Event Forwarding performance counter
                                    if (PerformanceCounterCategory.Exists(LogWatcher.EventForwardingCounterCategoryName))
                                    {
                                        PerformanceCounterCategory.Delete(LogWatcher.EventForwardingCounterCategoryName);
                                    }

                                }
                                catch(System.UnauthorizedAccessException )
                                {
                                    // need to do something useful here
                                }


                                Console.WriteLine("Deleted event source {0} in the Application log.", LogWatcher.ApplicationName);
                                break;
                            } // case "-u":
                        default:
                            {
                                Console.WriteLine("Interactive usage: Specify -i or -u for install or uninstall. Other parameters ignored.");
                                break;
                            } // default:
                    } // switch (args[0].ToLower())
                } // if (args.Length > 0)
                else
                {
                    // called without any command-line parameters
                    Console.WriteLine("Interactive usage: Specify -i or -u for install or uninstall. Other parameters ignored.");
                    Environment.Exit(0);
                }
            } // if (System.Environment.UserInteractive)
            else
            {
                // not an interactive session, start the service.
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] { new LogWatcherSvc() };
                ServiceBase.Run(ServicesToRun);
            } // else if (System.Environment.UserInteractive)
        } // static void Main()
    } // static class Program
} // namespace LogWatcherSvc
