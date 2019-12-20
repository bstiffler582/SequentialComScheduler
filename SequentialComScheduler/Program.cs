using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace SequentialComScheduler
{
    class Program
    {
        private static SequentialComScheduler _comScheduler = new SequentialComScheduler();

        static void Main(string[] args)
        {
            // using TaskFactory
            Task.Factory.StartNew(() => 
            {
                Console.WriteLine("Task 1 Started...");
                Console.WriteLine("Task Thread Apartment State: " + Thread.CurrentThread.GetApartmentState().ToString());
                Console.WriteLine("Thread ID: " + Thread.CurrentThread.ManagedThreadId);
                Thread.Sleep(2500);
                Console.WriteLine("Task 1 Complete!");
            }, CancellationToken.None, TaskCreationOptions.None, _comScheduler);

            // using a dedicated Task
            Task comTask = new Task(() => Task2());
            comTask.Start(_comScheduler);

            // Cancelable task
            CancellationTokenSource tokenSrc = new CancellationTokenSource();
            CancellationToken token = tokenSrc.Token;

            Task.Factory.StartNew(() =>
            {
                Console.WriteLine("Task 3 Started...");
                Console.WriteLine("Thread ID: " + Thread.CurrentThread.ManagedThreadId);
                Console.WriteLine("Press any key to Cancel...");

                for (int i = 0; i < 10000; i++)
                {
                    if (token.IsCancellationRequested)
                    {
                        Console.WriteLine("Task 3 Canceled!");
                        break;
                    }
                    // simulate some work
                    Thread.SpinWait(100000);
                }

                Console.WriteLine("Task 3 Complete!");
            }, token, TaskCreationOptions.None, _comScheduler);

            
            Console.Read();

            // cancel task
            tokenSrc.Cancel();

            Thread.Sleep(1000);

            // clean up
            _comScheduler.Dispose();
        }

        private static void Task2()
        {
            Console.WriteLine("Task 2 Started...");
            Console.WriteLine("Thread ID: " + Thread.CurrentThread.ManagedThreadId);
            Thread.Sleep(2500);
            Console.WriteLine("Task 2 Complete!");
        }
    }
}
