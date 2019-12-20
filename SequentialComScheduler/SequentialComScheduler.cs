using System;
using System.Collections.Generic;
using System.Threading;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace SequentialComScheduler
{
    /// <summary>
    /// TaskScheduler implementation that uses a dedicated STA thread,
    /// and implements an IOleMessageFilter interface to handle
    /// COM exceptions and blocking
    /// </summary>
    class SequentialComScheduler : TaskScheduler, IDisposable
    {
        readonly BlockingCollection<Task> _taskQueue = new BlockingCollection<Task>();
        readonly Thread _thread;
        readonly CancellationTokenSource _cancelTake = new CancellationTokenSource();
        volatile bool _terminate = false;

        /// <summary>
        /// Constructor
        /// Initializes and starts STA thread
        /// </summary>
        public SequentialComScheduler()
        {
            // initialize STA thread for handling tasks
            _thread = new Thread(Run);
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        void Run()
        {
            // loop until disposed
            while (!_terminate)
            {
                try
                {
                    // pop task from queue and execute
                    var task = _taskQueue.Take(_cancelTake.Token);

                    // Register COM message filter
                    MessageFilter.Register();

                    // Run task
                    TryExecuteTask(task);

                    // Revoke COM message filter
                    MessageFilter.Revoke();
                }
                catch (OperationCanceledException) { }
            }

            // cleanup queue when loop terminates
            _taskQueue.Dispose();
        }

        protected override IEnumerable<Task> GetScheduledTasks()
        {
            return _taskQueue;
        }

        protected override void QueueTask(Task task)
        {
            _taskQueue.Add(task);
        }

        protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            if (Thread.CurrentThread == _thread)
            {
                return TryExecuteTask(task);
            }
            return false;
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        // cancel active task, clear queue and clean up
        private void Dispose(bool disposing)
        {
            if (!disposing) return;
            if (_cancelTake != null && _cancelTake.Token.CanBeCanceled) _cancelTake.Cancel();
            _terminate = true;
            _taskQueue.CompleteAdding();
        }
    }

    /// <summary>
    /// Interface / Class for wrapping COM operations in STA Thread.
    /// Handles rejected COM calls caused by blocking, e.g.
    /// RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, etc.
    /// 
    /// </summary>
    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {

        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);


        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);


        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    public class MessageFilter : IOleMessageFilter
    {
        public static void Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            int test = CoRegisterMessageFilter(newFilter, out oldFilter);

            if (test != 0)
            {
                Console.WriteLine(string.Format("CoRegisterMessageFilter failed with error : {0}", test));
            }
        }


        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            int test = CoRegisterMessageFilter(null, out oldFilter);
        }


        int IOleMessageFilter.HandleInComingCall(int dwCallType, System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo)
        {
            //returns the flag SERVERCALL_ISHANDLED. 
            return 0;
        }


        int IOleMessageFilter.RetryRejectedCall(System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // Thread call was refused, try again. 
            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER. 
            {
                // retry thread call at once, if return value >=0 & 
                // <100. 
                return 99;
            }
            return -1;
        }


        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            //return flag PENDINGMSG_WAITDEFPROCESS. 
            return 2;
        }

        // implement IOleMessageFilter interface. 
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    }
}