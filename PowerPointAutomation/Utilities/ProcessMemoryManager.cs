 using System;
using System.Diagnostics;

namespace PowerPointAutomation.Utilities
{
    public static class ProcessMemoryManager
    {
        public static void SetupProcessMemory()
        {
            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                currentProcess.PriorityClass = ProcessPriorityClass.AboveNormal;
                
                GC.Collect(2, GCCollectionMode.Forced, true, true);
                
                currentProcess.MinWorkingSet = new IntPtr(204800);   // 200MB minimum
                currentProcess.MaxWorkingSet = new IntPtr(1048576000); // 1GB maximum

                Console.WriteLine("Process memory settings optimized for PowerPoint automation");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Failed to optimize process memory settings: {ex.Message}");
            }
        }
    }
}