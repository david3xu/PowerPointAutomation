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

        // Batch size threshold for automatic release
        private const int AutoReleaseBatchSize = 100;
        
        // Track the creation time of objects for more efficient release
        private static Dictionary<object, DateTime> objectCreationTimes = new Dictionary<object, DateTime>();
        
        // Flag to temporarily pause automatic COM object release
        private static bool isPaused = false;

        /// <summary>
        /// Pauses automatic COM object release during critical operations
        /// </summary>
        public static void PauseRelease()
        {
            isPaused = true;
            Console.WriteLine("COM object auto-release paused");
        }

        /// <summary>
        /// Resumes automatic COM object release
        /// </summary>
        public static void ResumeRelease()
        {
            isPaused = false;
            Console.WriteLine("COM object auto-release resumed");
        }

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
                    int refCount = Marshal.ReleaseComObject(obj);
                    // Uncomment for debugging
                    // Console.WriteLine($"Released COM object. Remaining references: {refCount}");
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
        /// Alternative version that doesn't use ref parameter (to fix compiler errors)
        /// </summary>
        /// <typeparam name="T">Type of COM object to release</typeparam>
        /// <param name="obj">The COM object to release</param>
        public static void ReleaseCOMObject<T>(T obj) where T : class
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
                objectCreationTimes[obj] = DateTime.Now;
                
                // Auto-release if we've accumulated too many objects (but only if not paused)
                if (!isPaused && trackedObjects.Count >= AutoReleaseBatchSize)
                {
                    Console.WriteLine($"Auto-releasing COM objects (count: {trackedObjects.Count})");
                    ReleaseOldestObjects(AutoReleaseBatchSize / 2); // Release half of the tracked objects
                }
            }
        }

        /// <summary>
        /// Releases the oldest tracked COM objects up to the specified count
        /// </summary>
        /// <param name="count">Number of oldest objects to release</param>
        public static void ReleaseOldestObjects(int count)
        {
            // If release is paused or count is 0, don't release anything
            if (isPaused || count <= 0)
            {
                return;
            }
            
            // Sort objects by creation time
            var objectsByAge = new List<object>(trackedObjects);
            objectsByAge.Sort((a, b) => objectCreationTimes[a].CompareTo(objectCreationTimes[b]));
            
            // Release the oldest objects
            int releaseCount = Math.Min(count, objectsByAge.Count);
            for (int i = 0; i < releaseCount; i++)
            {
                var obj = objectsByAge[i];
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
                finally
                {
                    trackedObjects.Remove(obj);
                    objectCreationTimes.Remove(obj);
                }
            }
            
            // Force garbage collection after releasing objects
            if (releaseCount > 0)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Releases all tracked COM objects
        /// </summary>
        /// <param name="batchSize">Process objects in batches of this size</param>
        public static void ReleaseAllTrackedObjects(int batchSize = 0)
        {
            Console.WriteLine($"Releasing all tracked COM objects (count: {trackedObjects.Count})");
            
            // If batch size is specified and valid, process in batches
            if (batchSize > 0 && trackedObjects.Count > batchSize)
            {
                int totalBatches = (trackedObjects.Count + batchSize - 1) / batchSize;
                Console.WriteLine($"Processing in {totalBatches} batches of {batchSize}");
                
                for (int batch = 0; batch < totalBatches; batch++)
                {
                    int batchStart = batch * batchSize;
                    int batchEnd = Math.Min(batchStart + batchSize, trackedObjects.Count);
                    int batchCount = batchEnd - batchStart;
                    
                    // Create a batch of objects to release
                    List<object> batchObjects = new List<object>(batchCount);
                    for (int i = batchStart; i < batchEnd; i++)
                    {
                        if (i < trackedObjects.Count)
                        {
                            batchObjects.Add(trackedObjects[i]);
                        }
                    }
                    
                    // Release objects in this batch
                    foreach (object obj in batchObjects)
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
                            Console.WriteLine($"Error releasing tracked COM object in batch: {ex.Message}");
                        }
                        finally
                        {
                            trackedObjects.Remove(obj);
                            objectCreationTimes.Remove(obj);
                        }
                    }
                    
                    // Force intermediate garbage collection between batches
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    
                    Console.WriteLine($"Released batch {batch + 1}/{totalBatches} - {batchObjects.Count} objects");
                }
            }
            else
            {
                // Process all objects at once (original behavior)
                foreach (object obj in new List<object>(trackedObjects))
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
                
                // Clear collections
            trackedObjects.Clear();
                objectCreationTimes.Clear();
            }

            // Force final garbage collection
            FinalCleanup();
        }

        /// <summary>
        /// Returns the number of currently tracked COM objects
        /// </summary>
        /// <returns>The count of tracked COM objects</returns>
        public static int GetTrackedObjectCount()
        {
            return trackedObjects.Count;
        }

        /// <summary>
        /// Forces garbage collection to clean up any lingering COM objects
        /// </summary>
        public static void FinalCleanup()
        {
            Console.WriteLine("Performing final cleanup with garbage collection");
            
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
        /// <param name="batchSize">Size of batches for releasing objects</param>
        public static void SafeExecute(Action action, int batchSize = 0)
        {
            try
            {
                action();
            }
            finally
            {
                ReleaseAllTrackedObjects(batchSize);
                FinalCleanup();
            }
        }

        /// <summary>
        /// Safely executes a function with COM objects and ensures cleanup
        /// </summary>
        /// <typeparam name="T">Return type of the function</typeparam>
        /// <param name="func">The function to execute</param>
        /// <param name="batchSize">Size of batches for releasing objects</param>
        /// <returns>The result of the function</returns>
        public static T SafeExecute<T>(Func<T> func, int batchSize = 0)
        {
            try
            {
                return func();
            }
            finally
            {
                ReleaseAllTrackedObjects(batchSize);
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