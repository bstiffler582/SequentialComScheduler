# SequentialComScheduler
A TaskScheduler implementation that runs synchronously on a STA thread.

This implementation of the System.Threading.Tasks.TaskScheduler class facilitates the use of the TaskFactory, async/await, etc. functionality on a dedicated STA thread.
It also implements an OLE/COM message filter which...

[3]"Enables handling of incoming and outgoing COM messages while waiting for responses from synchronous calls. You can use message filtering to prevent waiting on a synchronous call from blocking another application."

The scheduler was originally designed to be used for automating tasks in Visual Studio (using the EnvDTE namespace).

See Program.main() for example usage.

References:

TaskScheduler:
[1] https://docs.microsoft.com/en-us/dotnet/api/system.threading.tasks.taskscheduler?view=netframework-4.8

[2] https://www.journeyofcode.com/custom-taskscheduler-sta-threads/

IMessageFilter:
[3] https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.ole.interop.imessagefilter?view=visualstudiosdk-2019

[4] https://infosys.beckhoff.com/english.php?content=../content/1033/tc3_automationinterface/54043195771173899.html&id=

Visual Studio EnvDTE:
[5] https://docs.microsoft.com/en-us/dotnet/api/envdte?view=visualstudiosdk-2017
