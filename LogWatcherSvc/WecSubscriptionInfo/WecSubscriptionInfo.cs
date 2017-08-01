// <copyright file="WecSubscriptionInfo.cs" company="Microsoft Corporation">
//      Copyright (c) Microsoft Corporation.
// </copyright>
// <author>Ted Hardy</author>
// Using the registry keys that the Event Collector uses to store configuration and event source status, retrieves meta-data about subscriptions and host counts.
// some logic (noted in comments) mimics the actions of the event collector source code.
using System;
[assembly: CLSCompliant(true)]
namespace WecSubscription
{
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

    public class WecSubscriptionInfo : IDisposable
    {
        /// <summary>
        /// Object for the event collector top level registry key.
        /// </summary>
        private RegistryKey eventCollectorBasekey = null;

        /// <summary>
        /// Stores the global configuration mode heart beat values.
        /// </summary>
        private Dictionary<string, int> GlobalConfigurationModeHeartbeat = new Dictionary<string, int>();

        /// <summary>
        /// Path to the EventCollector registry key.
        /// </summary>
        private string baseRegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector";

        private object activeSourceLock = new object();

        private object totalSourceLock = new object();

        /// <summary>
        /// Initializes a new instance of the <see cref="WecSubscriptionInfo"/> object. Which encapsulates logic for retrieving Event Forwarding Subscription meta-data.
        /// </summary>
        public WecSubscriptionInfo()
        {
            if (!this.tryGetBaseRegistryConnection())
            {
                throw new InvalidOperationException("Unable to connect to registry for WEC subscription configuration.");
            }

            this.tryGetGlobalConfigurationHeartbeatValues();

        } // internal WecSubscriptionInfo()

        /// <summary>
        /// Connects to the top level registry key for the Event Collector
        /// </summary>
        /// <returns>True if successful, false if connection failed, key does not exist, or key is empty.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        private bool tryGetBaseRegistryConnection()
        {
            RegistryKey lmBase = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);

            if(null == lmBase)
            {
                return false;
            }

            // this will throw on access denied or key doesn't exist.
            this.eventCollectorBasekey = lmBase.OpenSubKey(this.baseRegistryPath);

            if (null == eventCollectorBasekey || eventCollectorBasekey.SubKeyCount == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        } // internal bool tryGetRegistryConnection()

        /// <summary>
        /// Enumerates the subscriptions event sources and determines the number of active sources and total sources.
        /// </summary>
        /// <param name="SubscriptionName">Name of the subscription to query</param>
        /// <param name="totalSources">total source count</param>
        /// <param name="activeSources">active source count</param>
        /// <returns>true if completed successfully</returns>
        public bool TryGetSubscriptionSourceCount(string SubscriptionName, out int totalSources, out int activeSources)
        {
            using (RegistryKey subKey = this.eventCollectorBasekey.OpenSubKey("Subscriptions\\" + SubscriptionName))
            {
                activeSources = 0;
                totalSources = 0;
                int subHeartbeatInterval = 0;

                if (null == subKey)
                {
                    // key (subscription) does not exist
                    totalSources = 0;
                    activeSources = 0;
                    return false;
                }

                string[] eventSources = null;

                // retrieve subscription source list
                if (this.tryGetSubscriptionSources(subKey, out eventSources))
                {
                    // get subscription heartbeat
                    if (this.tryGetSubscriptionHeartbeatInterval(subKey, out subHeartbeatInterval))
                    {
                        for (int i = 0; i < eventSources.Length; i++)
                        {
                            bool Active = false;
                            // check if source is active
                            if (this.tryGetSubscriptionSourceIsActive(subKey, eventSources[i], subHeartbeatInterval, out Active))
                            {
                                // check if returned as an active source.
                                if (Active)
                                {
                                    // increment active source list
                                    activeSources++;
                                }
                            }
                            // always increment total source list
                            totalSources++;
                        }
                    }
                    else
                    {
                        totalSources = 0;
                        activeSources = 0;
                        return false;
                    }
                }
                else
                {
                    totalSources = 0;
                    activeSources = 0;
                    return false;
                }
                return true;
            } // using (RegistryKey subKey = this.eventCollectorBasekey.OpenSubKey("Subscriptions\\" + SubscriptionName))
        } // public bool TryGetSubscriptionSourceCount(string SubscriptionName, out int totalSources, out int activeSources)

        /// <summary>
        /// Enumerates the subscriptions event sources and determines the number of active sources and total sources.
        /// </summary>
        /// <param name="SubscriptionName">Name of the subscription to query</param>
        /// <param name="totalSources">total source count</param>
        /// <param name="activeSources">active source count</param>
        /// <returns>true if completed successfully</returns>
        public bool TryGetSubscriptionSourceCount2(string SubscriptionName, out int totalSources, out int activeSources)
        {
            int localActiveSources = 0;
            int localTotalSources = 0;

            using (RegistryKey subKey = this.eventCollectorBasekey.OpenSubKey("Subscriptions\\" + SubscriptionName))
            {
                activeSources = 0;
                totalSources = 0;
                int subHeartbeatInterval = 0;

                if (null == subKey)
                {
                    // key (subscription) does not exist
                    totalSources = 0;
                    activeSources = 0;
                    return false;
                }

                string[] eventSources = null;

                // retrieve subscription source list
                if (this.tryGetSubscriptionSources(subKey, out eventSources))
                {
                    if (null != eventSources && eventSources.Length > 0)
                    {
                        // get subscription heartbeat
                        if (this.tryGetSubscriptionHeartbeatInterval(subKey, out subHeartbeatInterval))
                        {
                            ParallelLoopResult result = Parallel.ForEach(eventSources, (eventsource) =>
                            {
                                bool Active = false;
                                // check if source is active
                                if (this.tryGetSubscriptionSourceIsActive(
                                    subKey,
                                    eventsource,
                                    subHeartbeatInterval,
                                    out Active))
                                {
                                    // check if returned as an active source.
                                    if (Active)
                                    {
                                        // increment active source list
                                        lock (this.activeSourceLock)
                                        {
                                            localActiveSources++;
                                        }
                                    }
                                }
                                // always increment total source list
                                lock (this.totalSourceLock)
                                {
                                    localTotalSources++;
                                }
                            });

                            // crappy wait.
                            while(!result.IsCompleted)
                            {
                                System.Threading.Thread.Sleep(10);
                            }

                            totalSources = localTotalSources;
                            activeSources = localActiveSources;
                        }
                    }
                    else
                    {
                        totalSources = 0;
                        activeSources = 0;
                        return false;
                    }
                }
                else
                {
                    totalSources = 0;
                    activeSources = 0;
                    return false;
                }
                return true;
            } // using (RegistryKey subKey = this.eventCollectorBasekey.OpenSubKey("Subscriptions\\" + SubscriptionName))
        } // public bool TryGetSubscriptionSourceCount(string SubscriptionName, out int totalSources, out int activeSources)



        /// <summary>
        /// Enumerates subscriptions in registry and returns array of enabled subscription names.
        /// </summary>
        /// <param name="SubscriptionNames">populated with enabled subscriptions.</param>
        /// <returns>true if successfully completed, false if error encountered.</returns>
        public bool TryGetEnabledSubscriptions(out string[] SubscriptionNames)
        {
            using (RegistryKey subs = this.eventCollectorBasekey.OpenSubKey("Subscriptions"))
            {
                if (null == subs)
                {
                    // registry key for subscriptions not there - no subscriptions created yet.
                    SubscriptionNames = null;
                    return false;
                }

                List<string> enabledSubscriptions = new List<string>();
                string[] subscriptions = subs.GetSubKeyNames();

                if (null == subscriptions)
                {
                    // "subscriptions" registry key exists, but no subscriptions found - all subscriptions deleted.
                    SubscriptionNames = null;
                    return false;
                }

                // iterate over all subscriptions and check if enabled.
                for (int i = 0; i < subscriptions.Length; i++)
                {
                    using (RegistryKey subscriptionKey = subs.OpenSubKey(subscriptions[i]))
                    {
                        if (null == subscriptionKey)
                        {
                            continue;
                        }

                        bool subscriptionEnabled = false;

                        if (this.tryGetSubscriptionIsEnabled(
                            subscriptionKey, 
                            out subscriptionEnabled))
                        {
                            if (subscriptionEnabled)
                            {
                                // add to the list of enabled subscriptions.
                                enabledSubscriptions.Add(subscriptions[i]);
                            }
                        }
                    } // using(RegistryKey subscriptionKey = subs.OpenSubKey(subscriptions[i]))
                } // for(int i = 0; i < subscriptions.Length ;i++)

                SubscriptionNames = enabledSubscriptions.ToArray();
                return true;
            } // using(RegistryKey subs = this.eventCollectorBasekey.OpenSubKey("Subscriptions"))
        } // public bool TryGetEnabledSubscriptions(out string[] SubscriptionNames)

        /// <summary>
        /// Checks if the subscription is marked as enabled.
        /// </summary>
        /// <param name="SubscriptionKey">Registry object for the subscription</param>
        /// <param name="isEnabled">True if subscription is enabled</param>
        /// <returns>True if completed successfully.</returns>
        private bool tryGetSubscriptionIsEnabled(
            RegistryKey SubscriptionKey, 
            out bool isEnabled)
        {
            int enabledValue = 0;
            if (this.regReadDWORDValue(SubscriptionKey, "Enabled", out enabledValue))
            {
                if (enabledValue == 1)
                {
                    isEnabled = true;
                    return true;
                }
            }

            isEnabled = false;
            return false;
        } // private bool tryGetSubscriptionIsEnabled(

        /// <summary>
        /// Reads the specified double-word (integer) value from the supplied registry key.
        /// </summary>
        /// <param name="keyLocation">registry key location to query for value</param>
        /// <param name="valueName">name of value</param>
        /// <param name="value">output parameter of registry value data</param>
        /// <returns>True if successful, false if not exists or type mismatch</returns>
        private bool regReadDWORDValue(
            RegistryKey keyLocation, 
            string valueName, 
            out int value)
        {
            object regValue = keyLocation.GetValue(valueName);
            if (null == regValue)
            {
                value = 0;
                return false;
            }

            RegistryValueKind regValueKind = keyLocation.GetValueKind(valueName);
            if (regValueKind != RegistryValueKind.DWord)
            {
                value = 0;
                return false;
            }

            value = (int)regValue;
            return true;
        } // private bool regReadDWORDValue(

        /// <summary>
        /// Reads the specified quad-word value (long) from the supplied registry key.
        /// </summary>
        /// <param name="keyLocation">registry key location to query for value</param>
        /// <param name="valueName">name of value</param>
        /// <param name="value">output parameter of registry value data</param>
        /// <returns>True if successful, false if not exists or type mismatch</returns>
        private bool regReadQWORDValue(
            RegistryKey keyLocation, 
            string valueName, 
            out long value)
        {
            object regValue = keyLocation.GetValue(valueName);
            if (null == regValue)
            {
                // reg value not found.
                value = 0;
                return false;
            }

            RegistryValueKind regValueKind = keyLocation.GetValueKind(valueName);
            if (regValueKind != RegistryValueKind.QWord)
            {
                value = 0;
                return false;
            }

            value = (long)regValue;
            return true;
        } // private bool regReadQWORDValue(

        /// <summary>
        /// Reads the specified string value from the supplied registry key.
        /// </summary>
        /// <param name="keyLocation">registry key location to query for value</param>
        /// <param name="valueName">name of value</param>
        /// <param name="value">output parameter of registry value data</param>
        /// <returns>True if successful, false if not exists or type mismatch</returns>
        private bool regReadStringValue(
            RegistryKey keyLocation, 
            string valueName, 
            out string value)
        {
            object regValue = keyLocation.GetValue(valueName);
            if (null == regValue)
            {
                value = string.Empty;
                return false;
            }

            RegistryValueKind regValueKind = keyLocation.GetValueKind(valueName);
            if (regValueKind != RegistryValueKind.String)
            {
                value = string.Empty;
                return false;
            }

            value = (string)regValue;
            return true;
        } // private bool regReadStringValue(

        /// <summary>
        /// Reads the global configuration heartbeat intervals.
        /// </summary>
        /// <returns>True if completed successfully, falst if error occurred.</returns>
        private bool tryGetGlobalConfigurationHeartbeatValues()
        {
            using (RegistryKey configModes = this.eventCollectorBasekey.OpenSubKey("ConfigurationModes"))
            {
                // subkey does not exist,  return error.
                if (null == configModes)
                {
                    return false;
                }

                string[] globalConfigurationModes = configModes.GetSubKeyNames();

                if (null == globalConfigurationModes || globalConfigurationModes.Length == 0)
                {
                    return false;
                }

                for (int i = 0; i < globalConfigurationModes.Length; i++)
                {
                    int heartbeatInterval = 0;
                    using (RegistryKey mode = configModes.OpenSubKey(globalConfigurationModes[i]))
                    {
                        if (this.regReadDWORDValue(mode, "HeartbeatInterval", out heartbeatInterval))
                        {
                            if (this.GlobalConfigurationModeHeartbeat.ContainsKey(globalConfigurationModes[i]))
                            {
                                this.GlobalConfigurationModeHeartbeat[globalConfigurationModes[i]] = heartbeatInterval;
                            }
                            else
                            {
                                this.GlobalConfigurationModeHeartbeat.Add(globalConfigurationModes[i], heartbeatInterval);
                            }
                        }
                    } // using (RegistryKey mode = configModes.OpenSubKey(globalConfigurationModes[i]))
                } // for (int i = 0; i < globalConfigurationModes.Length; i++)
                return true;
            } // using (RegistryKey configModes = this.eventCollectorBasekey.OpenSubKey("ConfigurationModes"))
        } // private bool tryGetGlobalConfigurationHeartbeatValues()

        /// <summary>
        /// Returns the subscriptions heartbeat interval, whether custom set or from the Subscriptions configuration mode.
        /// </summary>
        /// <param name="SubscriptionKey">Registry object for the subscription</param>
        /// <param name="HeartbeatInterval">out parameter for the heartbeat interval</param>
        /// <returns>True if completed successfully, false if not.</returns>
        private bool tryGetSubscriptionHeartbeatInterval(
            RegistryKey SubscriptionKey, 
            out int HeartbeatInterval)
        {
            int customHeartbeatInterval = 0;
            // first check for a custom defined heartbeat interval
            if (this.regReadDWORDValue(
                SubscriptionKey, 
                "HeartbeatInterval", 
                out customHeartbeatInterval))
            {
                HeartbeatInterval = customHeartbeatInterval;
                return true;
            }
            else
            {
                // no custom heartbeat interval set, look up the configuration mode and return that value.
                string configMode = string.Empty;
                if (this.regReadStringValue(
                    SubscriptionKey, 
                    "ConfigurationMode", 
                    out configMode))
                {
                    if (this.GlobalConfigurationModeHeartbeat.TryGetValue(configMode, out HeartbeatInterval))
                    {
                        return true;
                    }
                    else
                    {
                        // config mode was not defined - this is an error/misconfig situation
                        HeartbeatInterval = 0;
                        return false;
                    }
                }
                else
                {
                    // no config mode defined - error mode
                    // this should never happen.
                    HeartbeatInterval = 0;
                    return false;
                }
            } // else this.regReadDWORDValue(
        } // private bool tryGetSubscriptionHeartbeatInterval(


        /// <summary>
        /// Enumerates the event sources that are under the subscription's event sources registry key.
        /// </summary>
        /// <param name="subscriptionKey"></param>
        /// <param name="subscriptionSources"></param>
        /// <returns>True if successful, false if EventSources key does not exist.</returns>
        private bool tryGetSubscriptionSources(RegistryKey subscriptionKey, out string[] subscriptionSources)
        {
            using (RegistryKey SourcesKey = subscriptionKey.OpenSubKey("EventSources"))
            {

                if (null == SourcesKey)
                {
                    //SourcesKey.Close();
                    // no EventSources subkey - misconfiguration.
                    subscriptionSources = null;
                    return false;
                }

                subscriptionSources = SourcesKey.GetSubKeyNames();
                return true;
            }
        }

        /// <summary>
        /// Querys the registry for a single WEF sources, using the last heartbeat time is calculates if this source would be active or not.
        /// </summary>
        /// <param name="subscriptionKey">registry key for the subscription</param>
        /// <param name="SubscriptionSource">Source host name in Fqdn format</param>
        /// <param name="heartbeatInterval">heartbeat interval for the subscription</param>
        /// <param name="isActive">True if source is active</param>
        /// <returns>true if completed without error</returns>
        private bool tryGetSubscriptionSourceIsActive(
            RegistryKey subscriptionKey, 
            string SubscriptionSource, 
            int heartbeatInterval, 
            out bool isActive)
        {
            using (RegistryKey SourceKey = subscriptionKey.OpenSubKey("EventSources\\" + SubscriptionSource))
            {
                long heartbeatFileTime = 0;
                int lastError = 0;

                if (null == SourceKey)
                {
                    isActive = false;
                    return false;
                }

                // Check for reported error (not always present)
                // if exists and is non-zero that means host is in error/not-active - regardless of value
                // matches WEC server logic.
                if (this.regReadDWORDValue(SourceKey, "LastError", out lastError))
                {
                    if (lastError != 0)
                    {
                        isActive = false;
                        return true;
                    }
                }

                if (this.regReadQWORDValue(SourceKey, "LastHeartbeatTime", out heartbeatFileTime))
                {
                    {
                        DateTime sourceLastHeartbeatTime = DateTime.FromFileTimeUtc(heartbeatFileTime);
                        DateTime subscriptionHeartbeatLimit = DateTime.Now.Subtract(TimeSpan.FromMilliseconds(heartbeatInterval));

                        if (sourceLastHeartbeatTime > subscriptionHeartbeatLimit)
                        {
                            isActive = true;
                            return true;
                        }
                        else
                        {
                            isActive = false;
                            return true;
                        }
                    }
                }
                else
                {
                    // matches WEC server logic - if lastHeartbeatFileTime key is not present then client is assumed active.
                    // it may have checked in initially to subscription but not yet sent a heartbeat signal.
                    isActive = true;
                    return true;
                }
            } // using (RegistryKey SourceKey = subscriptionKey.OpenSubKey("EventSources\\" + SubscriptionSource))
        } // private bool tryGetSubscriptionSourceIsActive(

        protected virtual void Dispose(bool CleanManaged)
        {
            this.closeRegistryKey();

            if(CleanManaged)
            {
                this.GlobalConfigurationModeHeartbeat.Clear();
            }
            return;
        }

        /// <summary>
        /// Disposes of resources and closes registry key links.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void closeRegistryKey()
        {
            if (this.eventCollectorBasekey != null)
            {
                this.eventCollectorBasekey.Close();
            }
        }
    } // class WecSubscriptionInfo
} // namespace WecSubscription
