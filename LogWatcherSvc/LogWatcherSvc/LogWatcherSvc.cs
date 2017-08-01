using System;
using System.Collections.Generic;
using System.Configuration;
using System.Configuration.Install;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Threading.Tasks;
using WecSubscription;

namespace LogWatcherSvc
{
    /// <summary>
    /// Container for the performance counters for the Event Collector category
    /// </summary>
    internal struct wefPerformanceCounters
    {
        /// <summary>
        /// Performance counter object for currently active event sources.
        /// </summary>
        internal PerformanceCounter ActiveSources;

        /// <summary>
        /// Performance counter object for total event sources (active + inactive.)
        /// </summary>
        internal PerformanceCounter TotalSources;
    }


    public partial class LogWatcherSvc : ServiceBase
    {
        /// <summary>
        /// (app.config) List of valid channels configured for monitoring.
        /// </summary>
        internal static List<string> ChannelNames = new List<string>();

        /// <summary>
        /// (app.config) interval in milliseconds that the event channels will be sampled.
        /// </summary>
        internal static double EpsTimerInterval = 1000;

        /// <summary>
        /// collection of performance counters, one for each channel.
        /// </summary>
        internal static List<PerformanceCounter> channelPerfCounterInstances = new List<PerformanceCounter>();

        internal static Dictionary<string, wefPerformanceCounters> wefPerformanceCounterInstances = new Dictionary<string, wefPerformanceCounters>();

        /// <summary>
        /// Collection of last known EventRecord ID values
        /// </summary>
        internal static List<ulong> ChannelEventRecordID = new List<ulong>();

        /// <summary>
        /// Collection of event log query objects, one for each channel.
        /// </summary>
        internal static List<EventLogQuery> ChannelQuery = new List<EventLogQuery>();

        /// <summary>
        /// flag if the timer has fired for the first time or not. First time will not publish EPS rate to performance counter.
        /// </summary>
        internal static bool FirstRun = true;

        /// <summary>
        /// global connection to the local event log.
        /// </summary>
        internal static EventLogSession EvtLogConnection = null;

        /// <summary>
        /// flag is the elapsed timer method is currently running.
        /// </summary>
        internal static bool EpsCounterRunning = false;

        /// <summary>
        /// displays EPS each interval when elapses.
        /// </summary>
        internal static Timer _epsTimer = null;

        /// <summary>
        /// flag if the method for getting subscription information is running. 
        /// </summary>
        private static bool subscriptionEnumRunning = false;

        public LogWatcherSvc()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Called by service control manager. Initializes and starts timer for checking event channels.
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            System.Diagnostics.EventLog.WriteEntry(
                LogWatcher.ApplicationName,
                "Application Starting up.",
                EventLogEntryType.Information,
                1000
                );

            // this is the "Service Start" phase, so it should return quickly.

            // connect to the event log service
            LogWatcherSvc.EvtLogConnection = EventLogSession.GlobalSession;

            // Parse App.config file
            #region Read App.Config File settings and create Performance Counter Instances for each valid channel

            try
            {
                string tmpChannelNames = System.Configuration.ConfigurationManager.AppSettings["EventChannel"].ToString();
                // Check if Channel exists locally
                string[] ChannelArray = tmpChannelNames.Split(';');

                for(int i = 0; i < ChannelArray.Length; i++)
                {
                    string ChannelToTest = ChannelArray[i].Trim();

                    // performance coutner object for the channel
                    PerformanceCounter PerfCounter = null;

                    // scrubbed (illegal characters mapped to legal ones) for Performance Channel instance name.
                    string ScrubbedChannelName = string.Empty;

                    // skip empty channel names (mostly for trailing ; characters)
                    if (ChannelToTest.Length > 0)
                    {
                        // attempt to initialize an EventLog object for that channel name. if it throws an exception the log is invalid/doesn't exist.
                        try
                        {
                            // if the event channel name does not exist, then the following will throw an EventLogNotFoundException
                            EventLogReader ReaderTest = new EventLogReader(ChannelToTest, PathType.LogName);
                            // this will test ability to access a channel
                            ReaderTest.ReadEvent();

                            // check that it is not a duplicate entry
                            if (!LogWatcherSvc.ChannelNames.Contains(ChannelToTest))
                            {
                                LogWatcherSvc.ChannelNames.Add(ChannelToTest);

                                System.Diagnostics.EventLog.WriteEntry(
                                    LogWatcher.ApplicationName,
                                    string.Format(
                                        "Reading app.config: added {0} event channel to active monitoring list", 
                                        ChannelToTest),
                                    EventLogEntryType.Information,
                                    1002
                                    );

                                ScrubbedChannelName = LogWatcherSvc.scrubInstanceName(ChannelToTest);

                                if (ScrubbedChannelName.CompareTo(ChannelToTest) != 0)
                                {
                                    // if there were changes made, log an entry for the admin.
                                    System.Diagnostics.EventLog.WriteEntry(
                                        LogWatcher.ApplicationName,
                                        string.Format(
                                            "Reading app.config: Channel Name {0} contains illegal characters for Performance Instance name. Performance counter Instance name for this channel will be:{1}", 
                                            ChannelToTest, 
                                            ScrubbedChannelName),
                                        EventLogEntryType.Information,
                                        1003
                                        );
                                } // if (ScrubbedChannelName.CompareTo(ChannelToTest) != 0)

                                // connect to the performance counter named instance in read/write mode
                                // using the scrubbed name.
                                PerfCounter = new PerformanceCounter(
                                    LogWatcher.PerformanceCounterCategoryName,
                                    LogWatcher.PerformanceCounterName,
                                    ScrubbedChannelName,
                                    false);

                                LogWatcherSvc.channelPerfCounterInstances.Add(PerfCounter);

                                // set the initial EPS rate to zero.
                                LogWatcherSvc.ChannelEventRecordID.Add(0);

                                EventLogQuery QueryObject = new EventLogQuery(ChannelToTest, PathType.LogName);

                                // read starting with the newest event in the channel.
                                QueryObject.ReverseDirection = true;
                                LogWatcherSvc.ChannelQuery.Add(QueryObject);
                            } // if (!LogWatcherSvc.PerfCounterInstances.ContainsKey(ChannelToTest))
                            else
                            {
                                System.Diagnostics.EventLog.WriteEntry(
                                    LogWatcher.ApplicationName,
                                    string.Format("Reading app.config: {0} event channel is duplicated and is omitted from active list", ChannelToTest),
                                    EventLogEntryType.Warning,
                                    1004
                                    );
                            } // else if (!LogWatcherSvc.PerfCounterInstances.ContainsKey(ChannelToTest))
                        } // try
                        catch (EventLogNotFoundException)
                        {
                            System.Diagnostics.EventLog.WriteEntry(
                                LogWatcher.ApplicationName,
                                string.Format(
                                    "Reading app.config: {0} event channel does not exist on the local machine. Channel cannot be monitored and is omitted from active list", 
                                    ChannelToTest),
                                EventLogEntryType.Warning,
                                1005
                                );
                        }
                        catch( EventLogReadingException ELRex)
                        {
                            System.Diagnostics.EventLog.WriteEntry(
                                LogWatcher.ApplicationName,
                                string.Format(
                                    "Reading app.config: {0} event channel exists on the local machine but error encountered reading events.\nChannel cannot be monitored and is omitted from active list\nMessage:{1}",
                                    ChannelToTest,
                                    ELRex.Message),
                                EventLogEntryType.Warning,
                                1006
                                );
                        }
                    } // if (ChannelToTest.Length > 0)
                } // for(int i = 0; i < ChannelArray.Length; i++)

            } // try
            catch (ConfigurationErrorsException)
            {
                System.Diagnostics.EventLog.WriteEntry(
                    LogWatcher.ApplicationName,
                    string.Format("EventChannel not defined in the app.config. Exiting."),
                    EventLogEntryType.Error,
                    1001
                    );
                Environment.Exit(1);
            }

            try
            {
                // read interval value from app.config.
                string tmpEpsTimerInterval = System.Configuration.ConfigurationManager.AppSettings["EpsTimerInterval"].ToString();

                if (!double.TryParse(tmpEpsTimerInterval, out LogWatcherSvc.EpsTimerInterval))
                {
                    System.Diagnostics.EventLog.WriteEntry(
                        LogWatcher.ApplicationName,
                        string.Format("EpsTimerInterval is not a valid double value. Exiting."),
                        EventLogEntryType.Error,
                        1002
                        );
                    Environment.Exit(1);
                }

                // dumbass protection
                if (LogWatcherSvc.EpsTimerInterval < 1000)
                {
                    System.Diagnostics.EventLog.WriteEntry(
                        LogWatcher.ApplicationName,
                        string.Format("EpsTimerInterval has a minimum value of 1000 (meaning 1 second intervals). Exiting."),
                        EventLogEntryType.Error,
                        1003
                        );

                    Environment.Exit(1);
                }
            } // try
            catch (ConfigurationErrorsException)
            {
                System.Diagnostics.EventLog.WriteEntry(
                    LogWatcher.ApplicationName,
                    string.Format("EpsTimerInterval not defined in the app.config. Exiting."),
                    EventLogEntryType.Error,
                    1001
                    );

                Environment.Exit(1);
            }
            #endregion

            // Create performance counter instances for each active subscription
            using(WecSubscription.WecSubscriptionInfo subInfo = new WecSubscription.WecSubscriptionInfo())
            {
                string[] enabledSubscriptions = null;

                if(subInfo.TryGetEnabledSubscriptions(out enabledSubscriptions))
                {
                    if (null != enabledSubscriptions)
                    {
                        for(int i = 0; i < enabledSubscriptions.Length; i++)
                        {
                            wefPerformanceCounters subPerfInstances = new wefPerformanceCounters();

                            // create performance counters for active and total
                            subPerfInstances.ActiveSources = new PerformanceCounter(
                                LogWatcher.EventForwardingCounterCategoryName,
                                LogWatcher.EventForwardingCounterName[0], // active
                                enabledSubscriptions[i],
                                false);

                            subPerfInstances.TotalSources = new PerformanceCounter(
                                LogWatcher.EventForwardingCounterCategoryName,
                                LogWatcher.EventForwardingCounterName[1], // total
                                enabledSubscriptions[i],
                                false);

                            // add performance counters to storage dictionary.
                            if (LogWatcherSvc.wefPerformanceCounterInstances.ContainsKey(enabledSubscriptions[i]))
                            {
                                // entry already exists, update it
                                LogWatcherSvc.wefPerformanceCounterInstances[enabledSubscriptions[i]] = subPerfInstances;
                            }
                            else
                            {
                                // add a new entry to the dictionary.
                                LogWatcherSvc.wefPerformanceCounterInstances.Add(
                                    enabledSubscriptions[i],
                                    subPerfInstances);
                            }
                        } // for(int i = 0; i < enabledSubscriptions.Length; i++)
                    } // if (null != enabledSubscriptions)
                } // if(subInfo.TryGetActiveSubscriptions(out enabledSubscriptions))
            } // using(WecSubscription.WecSubscriptionInfo subInfo = new WecSubscription.WecSubscriptionInfo())

            // start timer objects

            // create timer object and set elapse interval.
            LogWatcherSvc._epsTimer = new Timer(LogWatcherSvc.EpsTimerInterval);

            // set timer elapsed action
            LogWatcherSvc._epsTimer.Elapsed += _epsTimer_Elapsed;
            LogWatcherSvc._epsTimer.Elapsed += _epsTimer_Elapsed2;

            // set timer to auto-reset afer elapsing, so timer keeps firing.
            LogWatcherSvc._epsTimer.AutoReset = true;

            // start the timer.
            LogWatcherSvc._epsTimer.Start();
        }

        /// <summary>
        /// Gets subscription population active/total source counts
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _epsTimer_Elapsed2(object sender, ElapsedEventArgs e)
        {
            if (!LogWatcherSvc.subscriptionEnumRunning)
            {
                try
                {
                    LogWatcherSvc.subscriptionEnumRunning = true;
                    using (WecSubscriptionInfo WecSubs = new WecSubscriptionInfo())
                    {
                        string[] enabledSubs = null;

                        if (WecSubs.TryGetEnabledSubscriptions(out enabledSubs))
                        {
                            ParallelLoopResult loops = Parallel.ForEach<string>(enabledSubs,
                                activeSubscription => LogWatcherSvc.populateEventCollectorValues(
                                    activeSubscription,
                                    WecSubs));

                            if (!loops.IsCompleted)
                            {
                                System.Threading.Thread.Sleep(10);
                            }
                        }
                        else
                        {
                            // nothing enabled, do nothing.
                        }
                    } // using(WecSubscriptionInfo WecSubs = new WecSubscriptionInfo())
                }
                catch(InvalidOperationException)
                {
                    // nada
                }
                finally
                {
                    LogWatcherSvc.subscriptionEnumRunning = false;
                }
            } // if (!LogWatcherSvc.subscriptionEnumRunning)
        } // void _epsTimer_Elapsed2(object sender, ElapsedEventArgs e)

        /// <summary>
        /// Removes illegal characters that are not allowed for a performance counter isntance from the supplied channel name.
        /// </summary>
        /// <param name="rawChannelName">Event Channel name to check</param>
        /// <returns>Channel name with invalid characters replaced with alternate characters.</returns>
        private static string scrubInstanceName(string rawChannelName)
        {
            string ChannelName = string.Empty;

            // Per: http://msdn.microsoft.com/en-us/library/system.diagnostics.performancecounter.instancename(v=vs.110).aspx
            // Instance names cannot contain the following characters: Do not use the characters "(", ")", "#", "\", or "/" in the instance name. 
            // check for existence of any of those and replace using mapping (see above page)
            if (rawChannelName.Contains('(')
                || rawChannelName.Contains(')')
                || rawChannelName.Contains('#')
                || rawChannelName.Contains('\\')
                || rawChannelName.Contains('/'))
            {
                // contains one or more of the illegal characters.
                // use the same mapping as the MSDN article above for mapping illegal->legal characters.
                ChannelName = rawChannelName.Replace('(', '[');
                ChannelName = ChannelName.Replace(')', ']');
                ChannelName = ChannelName.Replace('#', '-');
                ChannelName = ChannelName.Replace('\\', '-');
                ChannelName = ChannelName.Replace('/', '-');
                return ChannelName;
            }
            else
            {
                // nothing illegal, return same string.
                return rawChannelName;
            }
        } // private static string scrubInstanceName(string rawChannelName)

        /// <summary>
        /// Called when interval timer elapses. Checks target channels for latest event and log movement.
        /// </summary>
        /// <param name="sender">(ignored)</param>
        /// <param name="e">(ignored)</param>
        void _epsTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (!LogWatcherSvc.EpsCounterRunning)
            {
                LogWatcherSvc.EpsCounterRunning = true;

                // list of channels that encountered errors and will be removed from the active list.
                List<string> BadChannels = new List<string>();

                ulong CurrentEventRecordId = 0;
                ulong PreviousEventRecordId = 0;

                // for each Channel in ChannelEventRecordID
                // read EventRecordID from channel
                // store it in the ChannelEventRecordID dictionary
                for (int i = 0; i < LogWatcherSvc.ChannelNames.Count; i++)
                {
                    string ChannelName = LogWatcherSvc.ChannelNames[i];
                    CurrentEventRecordId = 0;
                    long? totalRecords = 0;

                    // get previous EventRecordID for this channel.
                    PreviousEventRecordId = LogWatcherSvc.ChannelEventRecordID[i];

                    using(EventLogSession logSession = new EventLogSession())
                    {
                        EventLogInformation logInfo = logSession.GetLogInformation(ChannelName, PathType.LogName);

                        totalRecords = logInfo.RecordCount.GetValueOrDefault(0);

                    } // using(EventLogSession logSession = new EventLogSession())

                    // only attempt to read channel if there are events. Otherwise ReadEvent will hang.
                    if (totalRecords > 0)
                    {
                        using (EventLogReader NewReader = new EventLogReader(LogWatcherSvc.ChannelQuery[i]))
                        {
                            try
                            {
                                // read a single event
                                using (EventRecord newEvent = NewReader.ReadEvent())
                                {
                                    // Get the new EventRecordID from the event read from the channel.
                                    CurrentEventRecordId = LogWatcherSvc.GetEventBookmarkID(newEvent.Bookmark);

                                #if TEDLOG
                                    System.Diagnostics.EventLog.WriteEntry(
                                        LogWatcher.ApplicationName,
                                        string.Format("Read Erid Value:{0} for channel:{1}.",
                                        CurrentEventRecordId,
                                        ChannelName),
                                        EventLogEntryType.Information,
                                        9000
                                        );
                                #endif
                                } // using (EventRecord newEvent = NewReader.ReadEvent())
                            }
                            catch (EventLogProviderDisabledException ElPDex)
                            {
                                System.Diagnostics.EventLog.WriteEntry(
                                    LogWatcher.ApplicationName,
                                    string.Format("EventLog Provider Disabled Exception thrown reading event channel {0} for EPS calculation.\nMessage:{0}\nRemoving channel from active list.", ElPDex.Message),
                                    EventLogEntryType.Error,
                                    2002
                                    );
                                BadChannels.Add(ChannelName);
                                continue;
                            }
                            catch (EventLogReadingException ElRex)
                            {
                                System.Diagnostics.EventLog.WriteEntry(
                                    LogWatcher.ApplicationName,
                                    string.Format("EventLog Reading Exception thrown reading event channel {0} for EPS calculation.\nMessage:{0}\nRemoving channel from active list.", ElRex.Message),
                                    EventLogEntryType.Error,
                                    2002
                                    );
                                BadChannels.Add(ChannelName);
                                continue;
                            }
                            catch (EventLogException ElEx)
                            {
                                System.Diagnostics.EventLog.WriteEntry(
                                    LogWatcher.ApplicationName,
                                    string.Format("EventLog Exception thrown reading event channel {0} for EPS calculation.\nMessage:{0}\nRemoving channel from active list.", ElEx.Message),
                                    EventLogEntryType.Error,
                                    2002
                                    );
                                BadChannels.Add(ChannelName);
                                continue;
                            }
                        } // using (EventLogReader NewReader = new EventLogReader(LogWatcherSvc.ChannelQuery[i]))
                    } // if (totalRecords > 0)
                    else
                    {
                    #if TEDLOG
                        System.Diagnostics.EventLog.WriteEntry(
                            LogWatcher.ApplicationName,
                            string.Format("Zero events present in channel:{0}.",
                            ChannelName),
                            EventLogEntryType.Information,
                            9001
                            );
                    #endif
                    }
                    // Persist the new EventRecordID value for the next time through.
                    LogWatcherSvc.ChannelEventRecordID[i] = CurrentEventRecordId;

                    // calculate EPS
                    long ChannelEps = 0;

                    ChannelEps = (long)Math.Floor(((CurrentEventRecordId - PreviousEventRecordId) / LogWatcherSvc.EpsTimerInterval) * 1000.0);

                    #if TEDLOG
                    System.Diagnostics.EventLog.WriteEntry(
                            LogWatcher.ApplicationName,
                            string.Format("CurrentErid:{1}, PrevErid:{2} in channel:{0}. EPS:{3}",
                            ChannelName,
                            CurrentEventRecordId,
                            PreviousEventRecordId,
                            ChannelEps),
                            EventLogEntryType.Information,
                            9002
                            );
                    #endif
                    // if NOT the first time through, update the performance counter
                    if (LogWatcherSvc.FirstRun == false)
                    {
                        // publish performance counter EPS data.
                        LogWatcherSvc.channelPerfCounterInstances[i].RawValue = ChannelEps;
                    } // if (!LogWatcherSvc.FirstRun)
                } // for (int i = 0; i < LogWatcherSvc.ChannelNames.Count; i++)

                LogWatcherSvc.FirstRun = false;

                if (BadChannels.Count > 0)
                {
                    // remove any channels that had errors above to prevent them from causing repeated errors in the future.
                    foreach(string BadChannel in BadChannels)
                    {
                        LogWatcherSvc.ChannelNames.Remove(BadChannel);
                    } // foreach(string BadChannel in BadChannels)
                } // if (BadChannels.Count > 0)
                LogWatcherSvc.EpsCounterRunning = false;
            } // if (!LogWatcherSvc.TimerElapsedRunning)
        } // void _epsTimer_Elapsed(object sender, ElapsedEventArgs e)

        private static void populateEventCollectorValues(
            string subscriptionName,
            WecSubscriptionInfo wecSubInfo)
        {
            int totalSources = 0;
            int activeSources = 0;
            if(wecSubInfo.TryGetSubscriptionSourceCount(subscriptionName,
                out totalSources,
                out activeSources))
            {
                LogWatcherSvc.wefPerformanceCounterInstances[subscriptionName].ActiveSources.RawValue = activeSources;
                LogWatcherSvc.wefPerformanceCounterInstances[subscriptionName].TotalSources.RawValue = totalSources;
            }
            else
            {
                LogWatcherSvc.wefPerformanceCounterInstances[subscriptionName].ActiveSources.RawValue = 0;
                LogWatcherSvc.wefPerformanceCounterInstances[subscriptionName].TotalSources.RawValue = 0;
            }
        } // private static void populateEventCollectorValues(

        /// <summary>
        /// Called by Service Control manager. Disables interval timer and removes counter instances.
        /// </summary>
        protected override void OnStop()
        {
            // this is the service stopping phase, so it should also return quickly.
            System.Diagnostics.EventLog.WriteEntry(
                LogWatcher.ApplicationName,
                string.Format("Shutdown signal received. Stopping timers."),
                EventLogEntryType.Information,
                3000
                );

                // stop timer objects
                LogWatcherSvc._epsTimer.Enabled = false;

            System.Diagnostics.EventLog.WriteEntry(
                LogWatcher.ApplicationName,
                string.Format("Removing Performance Counter instances."),
                EventLogEntryType.Information,
                3001
                );

                for (int i = 0; i < LogWatcherSvc.channelPerfCounterInstances.Count; i++)
                {
                    LogWatcherSvc.channelPerfCounterInstances[i].RawValue = 0;
                    LogWatcherSvc.channelPerfCounterInstances[i].RemoveInstance();
                }

            System.Diagnostics.EventLog.WriteEntry(
                LogWatcher.ApplicationName,
                string.Format("Shutdown completed. Service Exiting."),
                EventLogEntryType.Information,
                3002
                );
        } // protected override void OnStop()

        /// <summary>
        /// Grovels the memory structure of the EventBookmark object to retrieve the underlying channel Event Record Id value.
        /// </summary>
        /// <param name="evtBookmark">Event Bookmark object to parse</param>
        /// <returns>Channel Event Record Id of the bookmarked Event</returns>
        private static ulong GetEventBookmarkID(EventBookmark evtBookmark)
        {
            ulong EventRecordId = 0;
            StringBuilder BookmarkText = new StringBuilder();
            string[] BookmarkElements = null;
            byte[] BookmarkData = null;
            // this is some ugly grovelling of the Bookmark object - there's no public access for the EventRecordId member

            // short cut if null is passed in.
            if(null == evtBookmark)
            {
                return 0;
            }
            
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter BinFormat = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();

            // read the raw memory of the bookmark object.
            using (System.IO.MemoryStream BookmarkMemory = new System.IO.MemoryStream())
            {
                // read the object into the memory stream
                BinFormat.Serialize(BookmarkMemory, evtBookmark);

                // Just read the bytes directly of the memory stream
                BookmarkData = BookmarkMemory.ToArray();
            } // using (System.IO.MemoryStream mem = new System.IO.MemoryStream())

            // convert the bytes to ASCII (which appears to work well enough but is extremly brittle)
            string BookmarkString = BookmarkText.Append(System.Text.Encoding.ASCII.GetString(BookmarkData)).ToString();

            string tmpEventRecordID  = string.Empty;

            // pull out the RecordId value from the converted memory stream text.
            // use "RecordId" as an anchor and then look forward in the string to the next space character
            //string RecordID = BookmarkString.Substring(BookmarkString.IndexOf("RecordId"), (BookmarkString.IndexOf(" ", BookmarkString.IndexOf("RecordId")) - BookmarkString.IndexOf("RecordId")));

            // split the resulting string into an array.
            BookmarkElements = BookmarkString.Split(' ');

            // the bookmark elements can vary - depending on channel type and bookmark state. so check each array element looking for the RecordId string.
            for (int i = 0; i < BookmarkElements.Length; i++ )
            {
                if (BookmarkElements[i].Trim().StartsWith("RecordId='"))
                {
                    // found the RecordId value, strip out the surrounding text.
                    tmpEventRecordID = BookmarkElements[i].Replace("RecordId='", string.Empty).Replace("'", string.Empty);
                } // if (BookmarkElements[i].StartsWith("RecordId='"))
            } // for (int i = 0; i < BookmarkElements.Length; i++ )

            // try to parse the found eventRecordID, if it fails return 0 and log a warning.
            if (ulong.TryParse(tmpEventRecordID, out EventRecordId))
            {
                return EventRecordId;
            } // if (ulong.TryParse(tmpEventRecordID, out EventRecordId))
            else
            {
                System.Diagnostics.EventLog.WriteEntry(
                    LogWatcher.ApplicationName,
                    string.Format("Error parsing out EventRecordId value from Bookmark object. Attempted to parse:{0}", tmpEventRecordID),
                    EventLogEntryType.Warning,
                    2002
                    );
                return 0;
            } // else if (ulong.TryParse(tmpEventRecordID, out EventRecordId))
            /*
             * example Bookmark object array elements.
                            string[] foobar = BookmarkString.Split(' ');
            {string[9]}
                [0]: "\0\0\0\0????\0\0\0\0\0\0\0\f\0\0\0NSystem.Core,"
                [1]: "Version=4.0.0.0,"
                [2]: "Culture=neutral,"
                [3]: "PublicKeyToken=b77a5c561934e089\0\0\00System.Diagnostics.Eventing.Reader.EventBookmark\0\0\0\fBookmarkText\0\0\0\0\0\0o<BookmarkList>\r\n"
                [4]: ""
                [5]: "<Bookmark"
                [6]: "Channel='ForwardedEvents'"
                [7]: "RecordId='2895256966'"
                [8]: "IsCurrent='true'/>\r\n</BookmarkList>\v"
            foobar[7]
            "RecordId='2895256966'"
            foobar[7]
            "RecordId='2895256966'"
            foobar[7].Replace("RecordId='","")
            "2895256966'"
            foobar[7].Replace("RecordId='","").Replace("'",'')
            Empty character literal
            foobar[7].Replace("RecordId='","").Replace("'",string.Empty)
            */

        } // private static ulong GetEventBookmarkID2(EventBookmark evtBookmark)
    } // public partial class LogWatcherSvc : ServiceBase





    [RunInstaller(true)]
    public sealed class MyProjectInstaller : Installer
    {
        public MyProjectInstaller()
        {
            ServiceInstaller serviceInstaller1;
            ServiceProcessInstaller processInstaller;

            // Instantiate installers for process and services.
            processInstaller = new ServiceProcessInstaller();
            serviceInstaller1 = new ServiceInstaller();

            // The services run under the system account.
            processInstaller.Account = ServiceAccount.LocalSystem;

            // The services are started manually.
            serviceInstaller1.StartType = ServiceStartMode.Automatic;

            // ServiceName must equal those on ServiceBase derived classes.
            // from MSDN: The ServiceName cannot be null or have zero length. Its maximum size is 256 characters. 
            //  It also cannot contain forward or backward slashes, '/' or '\', or characters from the ASCII character set with value less than decimal value 32.
            serviceInstaller1.ServiceName = LogWatcher.ApplicationName;

            // set service to depend upon EventLog service.
            serviceInstaller1.ServicesDependedOn = new string[] { "EventLog" };

            // How the service will be displayed in the Service Control Manager.
            serviceInstaller1.DisplayName = "Windows Event Log Watcher";

            // description for the service.
            serviceInstaller1.Description = "Publishes Performance counter values for specific Event Channels";


            // Add installers to collection. Order is not important.
            Installers.Add(serviceInstaller1);
            Installers.Add(processInstaller);
        } // public MyProjectInstaller()
    } //public class MyProjectInstaller : Installer

} // namespace LogWatcherSvc
