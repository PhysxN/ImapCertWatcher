using System;
using System.IO;
using ImapCertWatcher.Utils;

namespace ImapCertWatcher.Utils
{
    /// <summary>
    /// Вспомогательный класс для безопасной работы с асинхронными операциями и управлением ресурсами.
    /// </summary>
    public static class AsyncHelper
    {
        /// <summary>
        /// Выполняет асинхронную операцию с гарантированным таймаутом и обработкой ошибок.
        /// </summary>
        public static async System.Threading.Tasks.Task ExecuteWithTimeoutAsync(
            Func<System.Threading.CancellationToken, System.Threading.Tasks.Task> asyncOperation,
            TimeSpan timeout,
            Action<Exception> onError = null,
            string operationName = "Async Operation")
        {
            try
            {
                using (var cts = new System.Threading.CancellationTokenSource(timeout))
                {
                    await asyncOperation(cts.Token);
                }
            }
            catch (OperationCanceledException)
            {
                string errorMsg = $"{operationName} превысила максимальное время ({timeout.TotalSeconds:F0} сек)";
                var timeoutEx = new TimeoutException(errorMsg);
                onError?.Invoke(timeoutEx);
            }
            catch (Exception ex)
            {
                onError?.Invoke(ex);
            }
        }

        /// <summary>
        /// Выполняет синхронную операцию с безопасной обработкой исключений.
        /// </summary>
        public static void ExecuteSafeAsync(
            Action asyncAction,
            Action<Exception> onError = null,
            string operationName = "Operation")
        {
            try
            {
                asyncAction?.Invoke();
            }
            catch (Exception ex)
            {
                onError?.Invoke(ex);
            }
        }

        /// <summary>
        /// Пытается выполнить операцию с указанным количеством повторов при ошибке.
        /// </summary>
        public static async System.Threading.Tasks.Task<T> ExecuteWithRetryAsync<T>(
            Func<System.Threading.Tasks.Task<T>> asyncOperation,
            int maxRetries = 3,
            TimeSpan? delayBetweenRetries = null,
            Action<int, Exception> onRetry = null)
        {
            delayBetweenRetries = delayBetweenRetries ?? TimeSpan.FromSeconds(1);

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    return await asyncOperation();
                }
                catch (Exception ex) when (attempt < maxRetries)
                {
                    onRetry?.Invoke(attempt, ex);
                    await System.Threading.Tasks.Task.Delay(delayBetweenRetries.Value);
                }
            }

            // Последний раз выполняем без обработки ошибки
            return await asyncOperation();
        }
    }

    /// <summary>
    /// Вспомогательный класс для управления ресурсами и очистки.
    /// </summary>
    public static class ResourceHelper
    {
        /// <summary>
        /// Безопасно закрывает соединение и освобождает ресурсы.
        /// </summary>
        public static void SafeDisposeConnection(System.Data.Common.DbConnection conn)
        {
            try
            {
                if (conn != null)
                {
                    if (conn.State == System.Data.ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disposing connection: {ex.Message}");
            }
        }

        /// <summary>
        /// Безопасно удаляет файл, перехватывая исключения.
        /// </summary>
        public static bool SafeDeleteFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deleting file {filePath}: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Безопасно удаляет директорию, перехватывая исключения.
        /// </summary>
        public static bool SafeDeleteDirectory(string dirPath)
        {
            try
            {
                if (Directory.Exists(dirPath))
                {
                    Directory.Delete(dirPath, true);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deleting directory {dirPath}: {ex.Message}");
            }
            return false;
        }
    }
}
