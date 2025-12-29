using System;
using System.Threading;
using System.Threading.Tasks;

namespace ImapCertWatcher.Utils
{
    /// <summary>
    /// Безопасный асинхронный таймер, который правильно обрабатывает исключения
    /// и предотвращает перекрытие выполнения async задач.
    /// </summary>
    public class SafeAsyncTimer : IDisposable
    {
        private Timer _timer;
        private readonly Func<CancellationToken, Task> _asyncCallback;
        private readonly Action<Exception> _errorHandler;
        private bool _isExecuting = false;
        private readonly object _syncLock = new object();
        private bool _disposed = false;

        public SafeAsyncTimer(Func<CancellationToken, Task> asyncCallback, Action<Exception> errorHandler = null)
        {
            _asyncCallback = asyncCallback ?? throw new ArgumentNullException(nameof(asyncCallback));
            _errorHandler = errorHandler ?? (ex => System.Diagnostics.Debug.WriteLine($"SafeAsyncTimer error: {ex}"));
        }

        /// <summary>
        /// Запускает таймер с указанными интервалами.
        /// </summary>
        public void Start(TimeSpan initialDelay, TimeSpan period)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SafeAsyncTimer));

            _timer = new Timer(async _ => await ExecuteAsync(), null, initialDelay, period);
        }

        /// <summary>
        /// Запускает таймер с одинаковым начальным и периодическим интервалом.
        /// </summary>
        public void Start(TimeSpan interval)
        {
            Start(interval, interval);
        }

        private async Task ExecuteAsync()
        {
            lock (_syncLock)
            {
                if (_isExecuting)
                    return; // Предотвращаем перекрытие выполнения
                _isExecuting = true;
            }

            try
            {
                var cts = new CancellationTokenSource(TimeSpan.FromMinutes(10)); // Timeout 10 мин
                await _asyncCallback(cts.Token);
            }
            catch (OperationCanceledException)
            {
                _errorHandler?.Invoke(new TimeoutException("SafeAsyncTimer: async операция превысила timeout"));
            }
            catch (Exception ex)
            {
                _errorHandler?.Invoke(ex);
            }
            finally
            {
                lock (_syncLock)
                {
                    _isExecuting = false;
                }
            }
        }

        public void Stop()
        {
            _timer?.Change(Timeout.Infinite, Timeout.Infinite);
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            Stop();
            _timer?.Dispose();
            _disposed = true;
        }
    }
}
