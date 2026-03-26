using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using static System.Net.Mime.MediaTypeNames;

namespace TransferTargetBridge
{
    [ComVisible(true)]
    [Guid("F7A1B2C3-D4E5-4F67-8901-ABCDEF123456")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITransferTargetHelper
    {
        /// <summary>
        /// Shares email content to a specific app identified by appId.
        /// </summary>
        string ShareEmailToApp(string appId, string subject, string body, string htmlBody, long hwnd);
    }

    [ComVisible(true)]
    [Guid("A8B9C0D1-E2F3-4A5B-6C7D-8E9F0A1B2C3D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("TransferTargetBridge.Helper")]
    public class TransferTargetHelper : ITransferTargetHelper
    {
        public string ShareEmailToApp(string appId, string subject, string body, string htmlBody, long hwnd)
        {
            return RunSync(() => ShareEmailToAppAsync(appId, subject, body, htmlBody, (IntPtr)hwnd));
        }

        // ──────────────────────────────────────────────
        // Async implementation
        // ──────────────────────────────────────────────

        private static DataPackage BuildEmailDataPackage(string subject, string body, string htmlBody)
        {
            var dp = new DataPackage();
            dp.Properties.Title = subject;
            dp.Properties.Description = "Shared via TransferTarget API";

            if (!string.IsNullOrEmpty(body))
            {
                dp.SetText(body);
            }

            if (!string.IsNullOrEmpty(htmlBody))
            {
                dp.SetHtmlFormat(HtmlFormatHelper.CreateHtmlFormat(htmlBody));
            }

            return dp;
        }

        private static Windows.UI.WindowId WindowIdFromHwnd(IntPtr hwnd)
        {
            return new Windows.UI.WindowId { Value = (ulong)hwnd.ToInt64() };
        }

        private async Task<string> ShareEmailToAppAsync(
            string appId, string subject, string body, string htmlBody, IntPtr hwnd)
        {
            if (!Windows.Foundation.Metadata.ApiInformation.IsTypePresent(
                "Windows.ApplicationModel.DataTransfer.TransferTargetWatcher"))
            {
                return "ERROR: TransferTargetWatcher is not available on this version of Windows. " +
                        "Requires Windows 11 Build 26100.7015 or later.";
            }

            // Build a query where we ask at most one app, and it must be the one we want.
            // If the app is found and supports the DataPackage, it will be reported via watcher.Added
            // before the watcher.EnumerationCompleted.
            var dp = BuildEmailDataPackage(subject, body, htmlBody);
            var options = new TransferTargetDiscoveryOptions(dp.GetView())
            {
                MaxAppTargets = 1,
                AllowedTargetAppIds = [appId]
            };

            var watcher = TransferTarget.CreateWatcher(options);
            var done = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            TransferTarget? target = null;

            watcher.Added += (sender, args) => target = args.Target;
            watcher.EnumerationCompleted += (sender, args) => done.SetResult();
            watcher.Start();

            await done.Task;

            watcher.Stop();

            if (target == null)
            {
                return "FAILED: No matching target app found.";
            }

            var result = await watcher.TransferToAsync(target, WindowIdFromHwnd(hwnd));
            if (result.Succeeded)
            {
                return "SUCCESS";
            }
            return $"FAILED: Transfer failed: {result.ExtendedError?.Message}";
        }

        // ──────────────────────────────────────────────
        // Sync-over-async helper for COM/VBA callers
        // ──────────────────────────────────────────────

        private static string RunSync(Func<Task<string>> asyncFunc)
        {
            var prevCtx = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(null);
            try
            {
                return asyncFunc().GetAwaiter().GetResult();
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(prevCtx);
            }
        }
    }
}
