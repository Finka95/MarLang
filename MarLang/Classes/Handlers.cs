using System;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using Telegram.Bot.Types.InputFiles;

namespace MarLang
{
    class Handlers
    {
        public static List<IsDeletedClass> isDeletedList {get; set;}
        public static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            var handler = update.Type switch
            {
                UpdateType.Message            => BotOnMessageReceived(botClient, update.Message!),
                UpdateType.EditedMessage      => BotOnMessageReceived(botClient, update.EditedMessage!),
                UpdateType.PollAnswer         => PoolAnswerReceived(botClient, update.PollAnswer!),
                _                             => UnknownUpdateHandlerAsync(botClient, update)
            };

            try
            {
                await handler;
            }
            catch (Exception exception)
            {
                await HandleErrorAsync(botClient, exception, cancellationToken);
                Console.WriteLine("HandleUpdateAsync exception");
            }
        }
        public static async Task PoolAnswerReceived(ITelegramBotClient botClient, PollAnswer answer)
        {
            if (!Excel.ListUsers.ContainsKey(answer.User.Id))
                Excel.AddNewUser(answer.User.Id);
            int mainIndex;
            Excel.ListUsers.TryGetValue(answer.User.Id, out mainIndex);
            if(Excel.arrayUsersQuizWords[mainIndex] == null)
                return;
            if(Excel.arrayUsersQuizWords[mainIndex].Count != 0)
                await GetQuiz(botClient, answer.User.Id);
        }

        public static async Task BotOnMessageReceived(ITelegramBotClient botClient, Message message)
        {
            if (!Excel.ListUsers.ContainsKey(message.Chat.Id))
                Excel.AddNewUser(message.Chat.Id);
            int mainIndex;
            Excel.ListUsers.TryGetValue(message.Chat.Id, out mainIndex);
            Console.WriteLine($"User login: {message.From.Username}\nUser text message: {message.Text}");
            if (message.Text.Contains(@"/start"))
            {
                string text = @$"Hello!{Environment.NewLine}I’m a bot that could be your assistant in English. Write the word you need to clarify and I’ll write it down in your word sheet and tell you what it means. If you already have a few words you can check your knowledge in the /quiz. To stop it press /stopquiz. You can also look at the /last10 words or /delete any of them. Also you can download excel file with all your words by writing /sendexcel";
                await Program.client.SendTextMessageAsync(message.Chat.Id, text);
            }
            else if (message.Text.Contains(@"/sendexcel"))
            {
                string path = Excel.CreateExcelFile(message.Chat.Id);
                using (var stream = System.IO.File.Open(path, FileMode.Open))
                {
                    InputOnlineFile file = new InputOnlineFile(stream, "words.xlsx");
                    await botClient.SendDocumentAsync(message.Chat.Id,file);
                }
            }
            else if (message.Text.Contains(@"/stopquiz"))
            {
                Excel.arrayUsersQuizWords[mainIndex] = new List<List<string>>();
            }
            else if (message.Text.Contains(@"/delete"))
            {
                isDeletedList.SetValue(message.Chat.Id, true);
                await botClient.SendTextMessageAsync(message.Chat.Id, "Write the word to be deleted", replyMarkup: new ReplyKeyboardRemove());
            }
            else if(message.Text.Contains(@"/quiz"))
            {
                Excel.SetQuizWords(message.Chat.Id);
                await GetQuiz(botClient, message.Chat.Id);
            }
            else if (message.Text.Contains(@"/last10"))
            {
                await Program.client.SendTextMessageAsync(message.Chat.Id, Excel.GetLastTenWords(message.Chat.Id));
            }
            else if(Regex.IsMatch(message.Text, "[a-z]+", RegexOptions.IgnoreCase) && !Excel.CheckWordInBook(message.Chat.Id, message.Text) && !isDeletedList.ReturnValue(message.Chat.Id))
            {
                List<string> descriprion = WebClientClass.GetDescriptions(message.Text);
                if(descriprion.Count < 1)
                {
                    await Program.client.SendTextMessageAsync(message.Chat.Id, "I got nothing(");
                    return;
                }
                StringBuilder answer = new StringBuilder();
                answer.Append(message.Text + "\n");
                for (int i = 0; i < descriprion.Count; i++)
                {
                    answer.Append($"{i + 1}. {descriprion[i]}\n");
                }
                await Program.client.SendTextMessageAsync(message.Chat.Id, answer.ToString());
                Excel.AddNewWord(message.Chat.Id, message.Text, descriprion);
            }
            else if(Regex.IsMatch(message.Text, "[a-z]+", RegexOptions.IgnoreCase) && isDeletedList.ReturnValue(message.Chat.Id))
            {
                if(Excel.DeleteWord(message.Chat.Id, message.Text))
                    await Program.client.SendTextMessageAsync(message.Chat.Id, $"word \"{message.Text}\" deleted!");
                else
                    await Program.client.SendTextMessageAsync(message.Chat.Id, "I can’t find that word in your dictionary(");
                isDeletedList.SetValue(message.Chat.Id, false);
            }
        }

        public static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            var ErrorMessage = exception switch
            {
                ApiRequestException apiRequestException => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
                _ => exception.ToString()
            };

            Console.WriteLine(ErrorMessage);
            return Task.CompletedTask;
        }

        public static Task UnknownUpdateHandlerAsync(ITelegramBotClient botClient, Update update)
        {
            Console.WriteLine($"Unknown update type: {update.Type}");
            return Task.CompletedTask;
        }

        private static async Task GetQuiz(ITelegramBotClient botClient, long chatId)
        {
            QuizClass qc = Excel.GetQuizCard(chatId);
            await botClient.SendPollAsync(chatId, qc.Word, qc.Descriptions, false, PollType.Quiz, correctOptionId: qc.CorrectAnswer);
        }
    }
}