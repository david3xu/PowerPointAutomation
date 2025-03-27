using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PowerPointAutomation.Utilities
{
    /// <summary>
    /// Utility class for safely releasing COM objects to prevent memory leaks
    /// </summary>
    /// <remarks>
    /// When working with Office Interop, it's crucial to properly release COM objects
    /// to avoid memory leaks and orphaned processes. This class provides methods to 
    /// safely release COM objects and track them for batch release.
    /// </remarks>
    public static class ComReleaser
    {
        // Collection to track COM objects for batch release
        private static List<object> trackedObjects = new List<object>();

        /// <summary>
        /// Safely releases a COM object and sets the reference to null
        /// </summary>
        /// <param name="obj">Reference to the COM object to release</param>
        public static void ReleaseCOMObject(ref object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error releasing COM object: {ex.Message}");
                }
                finally
                {
                    obj = null;
                }
            }
        }

        /// <summary>
        /// Tracks a COM object for later batch release
        /// </summary>
        /// <param name="obj">COM object to track</param>
        public static void TrackObject(object obj)
        {
            if (obj != null)
            {
                trackedObjects.Add(obj);
            }
        }

        /// <summary>
        /// Releases all tracked COM objects
        /// </summary>
        public static void ReleaseAllTrackedObjects()
        {
            foreach (object obj in trackedObjects)
            {
                try
                {
                    if (obj != null)
                    {
                        Marshal.ReleaseComObject(obj);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error releasing tracked COM object: {ex.Message}");
                }
            }

            // Clear the list after releasing all objects
            trackedObjects.Clear();

            // Force garbage collection
            FinalCleanup();
        }

        /// <summary>
        /// Forces garbage collection to clean up any lingering COM objects
        /// </summary>
        public static void FinalCleanup()
        {
            // Run garbage collection twice to ensure all references are cleaned up
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Safely executes an action with COM objects and ensures cleanup
        /// </summary>
        /// <param name="action">The action to execute</param>
        public static void SafeExecute(Action action)
        {
            try
            {
                action();
            }
            finally
            {
                ReleaseAllTrackedObjects();
                FinalCleanup();
            }
        }

        /// <summary>
        /// Safely executes a function with COM objects and ensures cleanup
        /// </summary>
        /// <typeparam name="T">Return type of the function</typeparam>
        /// <param name="func">The function to execute</param>
        /// <returns>The result of the function</returns>
        public static T SafeExecute<T>(Func<T> func)
        {
            try
            {
                return func();
            }
            finally
            {
                ReleaseAllTrackedObjects();
                FinalCleanup();
            }
        }

        /// <summary>
        /// Checks if a process with the provided name is running
        /// </summary>
        /// <param name="processName">Name of the process to check</param>
        /// <returns>True if the process is running, false otherwise</returns>
        public static bool IsProcessRunning(string processName)
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(processName);
            return processes.Length > 0;
        }

        /// <summary>
        /// Attempts to kill all processes with the provided name
        /// </summary>
        /// <param name="processName">Name of the processes to kill</param>
        /// <returns>Number of processes killed</returns>
        public static int KillProcess(string processName)
        {
            int count = 0;
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(processName);

            foreach (System.Diagnostics.Process process in processes)
            {
                try
                {
                    process.Kill();
                    count++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error killing process {processName}: {ex.Message}");
                }
            }

            return count;
        }
    }
}