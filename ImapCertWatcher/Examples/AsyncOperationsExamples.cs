using System;
using System.Threading.Tasks;
using ImapCertWatcher.Utils;

namespace ImapCertWatcher.Examples
{
    /// <summary>
    /// Примеры использования новых утилит для безопасной работы с async операциями.
    /// </summary>
    public class AsyncOperationsExamples
    {
        // ==================== ПРИМЕР 1: SafeAsyncTimer ====================
        
        /// <summary>
        /// Пример использования SafeAsyncTimer вместо обычного Timer для периодических проверок.
        /// </summary>
        public void Example_SafeAsyncTimer()
        {
            // Создаем таймер с async callback
            var timer = new SafeAsyncTimer(
                asyncCallback: async cancellationToken =>
                {
                    // Ваша асинхронная операция здесь
                    await SomeAsyncOperation(cancellationToken);
                },
                errorHandler: ex => Console.WriteLine($"Ошибка в таймере: {ex.Message}")
            );

            // Запускаем с интервалом 5 минут
            timer.Start(TimeSpan.FromMinutes(5));

            // Позже можно остановить
            // timer.Stop();
            // timer.Dispose();
        }

        // ==================== ПРИМЕР 2: ExecuteWithTimeoutAsync ====================
        
        /// <summary>
        /// Пример выполнения операции с таймаутом (максимальное время ожидания).
        /// </summary>
        public async Task Example_ExecuteWithTimeoutAsync()
        {
            // Выполняем операцию с максимальным временем 30 секунд
            await AsyncHelper.ExecuteWithTimeoutAsync(
                asyncOperation: async cancellationToken =>
                {
                    // Вот здесь ваша длительная операция
                    // Если она займет > 30 сек, будет выброшено OperationCanceledException
                    await Task.Delay(5000, cancellationToken); // Имитируем долгую операцию
                },
                timeout: TimeSpan.FromSeconds(30),
                onError: ex => Console.WriteLine($"Ошибка при выполнении: {ex.Message}"),
                operationName: "Проверка почты"
            );
        }

        // ==================== ПРИМЕР 3: ExecuteWithRetryAsync ====================
        
        /// <summary>
        /// Пример повторного выполнения операции при ошибке (например, при нестабильном подключении).
        /// </summary>
        public async Task Example_ExecuteWithRetryAsync()
        {
            // Пытаемся получить данные, повторяя 3 раза при ошибке
            var result = await AsyncHelper.ExecuteWithRetryAsync(
                asyncOperation: async () =>
                {
                    // Ваша операция, которая может спонтанно сфейлиться
                    return await FetchDataFromApi();
                },
                maxRetries: 3,
                delayBetweenRetries: TimeSpan.FromSeconds(2),
                onRetry: (attempt, ex) =>
                {
                    Console.WriteLine($"Попытка {attempt} не удалась: {ex.Message}. Повтор через 2 сек...");
                }
            );
        }

        // ==================== ПРИМЕР 4: SafeDisposeConnection ====================
        
        /// <summary>
        /// Пример безопасного закрытия соединения с БД.
        /// </summary>
        public void Example_SafeDisposeConnection()
        {
            System.Data.Common.DbConnection conn = null;
            try
            {
                // Используете соединение...
                conn = new System.Data.SqlClient.SqlConnection("connection_string");
                conn.Open();
                // Выполняете запросы...
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
            finally
            {
                // Безопасно закрываем соединение, даже если случилась ошибка
                ResourceHelper.SafeDisposeConnection(conn);
            }
        }

        // ==================== ПРИМЕР 5: SafeDeleteFile ====================
        
        /// <summary>
        /// Пример безопасного удаления файла.
        /// </summary>
        public void Example_SafeDeleteFile()
        {
            string filePath = @"C:\Temp\myfile.txt";

            // Удаляем файл, но не выбрасываем исключение, если его нет
            bool deleted = ResourceHelper.SafeDeleteFile(filePath);
            
            if (deleted)
                Console.WriteLine("Файл успешно удален");
            else
                Console.WriteLine("Не удалось удалить файл (возможно, он не существует)");
        }

        // ==================== ПРИМЕР 6: Комбинированное использование ====================
        
        /// <summary>
        /// Пример использования нескольких утилит вместе для надежной операции.
        /// </summary>
        public async Task Example_CombinedUsage()
        {
            // Создаем таймер, который периодически проверяет почту с повторами при ошибке
            var timer = new SafeAsyncTimer(
                asyncCallback: async cancellationToken =>
                {
                    try
                    {
                        // Пытаемся получить письма 3 раза, ждем 2 сек между попытками
                        await AsyncHelper.ExecuteWithRetryAsync(
                            asyncOperation: async () =>
                            {
                                await DoMailCheck(cancellationToken);
                                return true;
                            },
                            maxRetries: 3,
                            delayBetweenRetries: TimeSpan.FromSeconds(2),
                            onRetry: (attempt, ex) =>
                            {
                                Console.WriteLine($"Попытка проверки почты #{attempt} не удалась, повтор...");
                            }
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Критическая ошибка при проверке почты: {ex.Message}");
                    }
                },
                errorHandler: ex => Console.WriteLine($"Ошибка таймера: {ex.Message}")
            );

            timer.Start(TimeSpan.FromMinutes(5)); // Каждые 5 минут

            // Имитируем работу
            await Task.Delay(TimeSpan.FromMinutes(10));

            timer.Dispose();
        }

        // ==================== Вспомогательные методы ====================

        private async Task SomeAsyncOperation(System.Threading.CancellationToken ct)
        {
            await Task.Delay(1000, ct);
        }

        private async Task<string> FetchDataFromApi()
        {
            // Имитируем API запрос
            await Task.Delay(500);
            return "Data from API";
        }

        private async Task DoMailCheck(System.Threading.CancellationToken ct)
        {
            // Проверяем почту
            await Task.Delay(2000, ct);
        }
    }
}
