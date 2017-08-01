
# LogWatcherSvc
Windows Event Log and WEC server monitoring Service. 

This runs as a windows service that reads specific windows event logs and Windows Event Collector registry keys to compute Performance counters related to each.

# Performance Counters
There are two new Performance Counters created:
## Windows Event Forwarding
## Windows Event Log

## Windows Event Forwarding
Under "Windows Event Forwarding" performance counter there are two performance objects created: Active Event Source Count and Total Event Source Count.

Under each performance object

Active Event Source Count is the number of machines that have sent events or performed a heartbeat check-in for that subscription

Total Event Source Count is the total number of machines that have connected to that subscription since it's creation (or the last time the registry keys have been groomed.)

## Windows Event Log
There is one Performance Counter Object "Channel EPS rate" - an instance for each event channel to be monitored is created.

# Install/Uninstall
Install: From an elevated command-prompt, run: LogWatcherSvc.exe -i

Uninstall: From an elevated command-prompt, run: LogWatcherSvc.exe -u

# Configuration
In the .Config file, edit the "EventChannel" value. This is a semi-colon delmited list of event channels to monitor and create an instance for in the "Windows Event Log / Channel EPS Rate" Performance Counter Object.
