using System;
using System.Threading;
using Telegram.Bot;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types.Enums;
using OfficeOpenXml;

namespace MarLang
{
    class Program
    {
        private static string token { get; set; } = "YOUR TOKEN";
        public static TelegramBotClient client;

        static void Main(string[] args)
        {
           Start();
        }

        static void Start()
        {
            client = new TelegramBotClient(token);
            using var cts = new CancellationTokenSource();
            ReceiverOptions receiverOptions = new ReceiverOptions() { AllowedUpdates = { } };
            client.StartReceiving(Handlers.HandleUpdateAsync, Handlers.HandleErrorAsync, receiverOptions, cts.Token);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Excel.GetListUsers();

            Console.WriteLine("Bot is running successfully");
            Console.ReadLine();

            cts.Cancel();
        }

        static async void SendMessage()
        {
            await client.SendPollAsync(475178131, "What the descriptions", new string[] {"first", "second", "third"}, false, PollType.Quiz, correctOptionId: 2);
        }
    }
}

