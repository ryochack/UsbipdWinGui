using Microsoft.Win32.TaskScheduler;
using System;
using System.Diagnostics;

namespace UsbipdGui
{
    internal static class StartupSettings
    {
        public static bool IsRegistered(in string appName)
        {
            using TaskService taskService = new();
            return taskService.GetTask(appName) is not null;
        }

        public static void RegisterTask(in string appName, in string exePath)
        {
            try
            {
                using TaskService taskService = new();
                if (taskService.GetTask(appName) is not null)
                {
                    // already registered
                    return;
                }
                TaskDefinition taskDefinition = taskService.NewTask();
                taskDefinition.RegistrationInfo.Description = appName;
                taskDefinition.Principal.RunLevel = TaskRunLevel.Highest;  // Run as administrator
                taskDefinition.Triggers.Add(new LogonTrigger());           // Run at startup
                taskDefinition.Actions.Add(new ExecAction(exePath));
                taskService.RootFolder.RegisterTaskDefinition(appName, taskDefinition);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Catch exception {ex}");
            }
        }

        public static void RemoveTask(in string appName)
        {
            try
            {
                using TaskService taskService = new();
                taskService.RootFolder.DeleteTask(appName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Catch exception {ex}");
            }
        }
    }
}
