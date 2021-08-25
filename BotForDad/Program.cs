using System;
using System.Text;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.Enums;
using System.IO;
using OfficeOpenXml;
using Telegram.Bot.Types.InputFiles;
using System.Collections.Generic;

namespace BotForDad
{
    class Program
    {
        private static readonly TelegramBotClient Bot = new TelegramBotClient("1946915196:AAH3zcNb-X8JMw70r0dEu6iNNElJWK-3PXk");
        public static ExcelPackage ReceivedPackage { get; set; }
        public static ExcelWorksheet ReceivedWorksheet { get; set; }
        public static FileInfo ReceivedFile { get; set; }
        public static FileInfo NewFile { get; set; }
        public static ExcelPackage NewPackage { get; set; }
        public static ExcelWorksheet NewWorksheet { get; set; }
        public static int Rows { get; set; }
        public static int Columns { get; set; }

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.OutputEncoding = Encoding.UTF8;
            Bot.OnMessage += Bot_OnMessage;
            Bot.OnMessageEdited += Bot_OnMessage;
            Bot.StartReceiving();
            Console.ReadLine();
        }

        private static async void Bot_OnMessage(object sender, MessageEventArgs e)
        {

            Console.WriteLine($"{DateTime.Now} the {e.Message.Chat.Username}({e.Message.Chat.Id}) Write: {e.Message.Text}");
            switch (e.Message.Text)
            {
                case "/start":
                    await Bot.SendTextMessageAsync(e.Message.Chat.Id, "🌸 Hi, please select one function that you need 🌸\n"
                        + "/deleteallexceptprice\n"
                        + "/deleteall\n"
                        + "/addpurchase\n");
                    break;

                case "/deleteallwithoutprice":
                    await Bot.SendTextMessageAsync(e.Message.Chat.Id, Commands.DeleteAllExceptPrice);
                    break;

                case "/deleteall":
                    await Bot.SendTextMessageAsync(e.Message.Chat.Id, Commands.DeleteAll);
                    break;

                case "/addpurchase":
                    await Bot.SendTextMessageAsync(e.Message.Chat.Id, Commands.AddPurchase);
                    break;
            }

            //deleteall
            if (e.Message.ReplyToMessage != null && e.Message.ReplyToMessage.Text.Contains(Commands.DeleteAll))
            {
                if (e.Message.Type == MessageType.Document)
                {
                    var mass = new List<string> { "Наименование", "Артикул", "数量",
                        "Кол-во", "型号", "总数量", "T.QTY总数量", "T.QTY", "DESCR NAME(CHINESE)", "型号ITEM NO" };
                    Delete(e, mass);
                }
            }

            //deleteallexceptprice
            if (e.Message.ReplyToMessage != null && e.Message.ReplyToMessage.Text.Contains(Commands.DeleteAllExceptPrice))
            {
                if (e.Message.Type == MessageType.Document)
                {
                    var mass = new List<string> { "Наименование", "Артикул", "数量", "Кол-во", "型号", "总数量", "T.QTY总数量", "T.QTY", "型号ITEM NO", "DESCR NAME(CHINESE)", "Цена", "Цена(ю.)", "PEICE单价" };
                    Delete(e, mass);
                }
            }

            //addpurchase
            if (e.Message.ReplyToMessage != null && e.Message.ReplyToMessage.Text.Contains(Commands.AddPurchase))
            {
                if (e.Message.Type == MessageType.Document) InsertTwo(e);
            }
        }

        public static async void Delete(MessageEventArgs e, List<string> mass)
        {
            var doc = await Bot.GetFileAsync(e.Message.Document.FileId);
            var name = $@"C:\Users\Кристина\Desktop\{e.Message.Document.FileName}";
            FileStream fs = new FileStream(name, FileMode.Create);
            await Bot.GetInfoAndDownloadFileAsync(doc.FileId, fs);

            ReceivedFile = new FileInfo(name);
            ReceivedPackage = new ExcelPackage(ReceivedFile);

            var path = @$"C:\Users\Кристина\Desktop\ResultWithDeletion_{DateTime.Now.ToShortTimeString()}.xlsx";
            NewFile = new FileInfo(path);
            NewPackage = new ExcelPackage(NewFile);
            for (int k = 0; k < ReceivedPackage.Workbook.Worksheets.Count; k++)
            {
                ReceivedWorksheet = ReceivedPackage.Workbook.Worksheets[k];
                Columns = ReceivedWorksheet.Dimension?.Columns ?? 0;
                NewWorksheet = NewPackage.Workbook.Worksheets.Add(ReceivedWorksheet.Name);
                await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);

                Rows = ReceivedWorksheet.Dimension?.Rows ?? 0;
                Console.WriteLine($"Sheet {ReceivedWorksheet.Name}");
                var currentColumn = 1;
                for (int j = 1; j <= Columns; j++)
                {
                    Console.WriteLine($"Column {j}");
                    var value = ReceivedWorksheet.Cells[1, j].Value?.ToString();
                    //if (mass.Contains(value) && value != null)
                    //{
                    //    for (int i = 1; i < Rows; i++)
                    //    {
                    //        NewWorksheet.Cells[i, currentColumn].Value = ReceivedWorksheet.Cells[i, j].Value;
                    //        await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);
                    //    }
                    //    currentColumn++;
                    //}
                    if (!mass.Contains(value))
                    {
                        NewWorksheet.DeleteColumn(j);
                        for (int i = 1; i < Rows; i++)
                        {
                            NewWorksheet.Cells[i, currentColumn].Value = ReceivedWorksheet.Cells[i, j].Value;
                            await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);

                        }
                    }
                }
                NewWorksheet.Cells["A1:G200"].AutoFitColumns();
                await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);

                for (int j = 1; j <= Columns; j++)
                {
                    for (int i = 1; i < Rows; i++)
                    {
                        if (NewWorksheet.Cells[i, j].Value?.ToString() == "Итого")
                        {
                            NewWorksheet.DeleteRow(i);
                            await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);
                        }
                    }
                }

                using var stream = File.OpenRead(path);
                InputOnlineFile iof = new InputOnlineFile(stream)
                {
                    FileName = NewFile.Name
                };
                await Bot.SendDocumentAsync(e.Message.Chat.Id, iof, "File");
            }
        }
        public static async void Insert(MessageEventArgs e)
        {
            var doc = await Bot.GetFileAsync(e.Message.Document.FileId);
            var name = $@"C:\Users\Кристина\Desktop\{e.Message.Document.FileName}";
            FileStream fs = new FileStream(name, FileMode.Create);
            await Bot.GetInfoAndDownloadFileAsync(doc.FileId, fs);

            ReceivedFile = new FileInfo(name);
            ReceivedPackage = new ExcelPackage(ReceivedFile);

            for (int k = 0; k < ReceivedPackage.Workbook.Worksheets.Count; k++)
            {
                ReceivedWorksheet = ReceivedPackage.Workbook.Worksheets[k];
                Columns = ReceivedWorksheet.Dimension?.Columns ?? 0;
                Rows = ReceivedWorksheet.Dimension?.Rows ?? 0;

                ReceivedWorksheet.InsertColumn(Columns + 1, 1);
                Columns = ReceivedWorksheet.Dimension?.Columns ?? 0;

                ReceivedWorksheet.Cells[1, Columns].Value = "Закупка";
                await ReceivedPackage.SaveAsAsync(new FileInfo(name)).ConfigureAwait(false);
                int index = 0;
                for (int j = 1; j <= Columns; j++)
                {
                    if (ReceivedWorksheet.Cells[1, j].Value?.ToString() == "Цена" ||
                        ReceivedWorksheet.Cells[1, j].Value?.ToString() == "Цена(ю.)")
                    {
                        index = j;
                    }
                }
                for (int i = 2; i <= Rows; i++)
                {
                    ReceivedWorksheet.Cells[i, Columns].Value = (((IConvertible)ReceivedWorksheet.Cells[i, index].Value).ToDouble(null) * 6.3 / 1.3).ToString();
                    await ReceivedPackage.SaveAsAsync(new FileInfo(name)).ConfigureAwait(false);
                }
                using var stream = File.OpenRead(name);
                InputOnlineFile iof = new InputOnlineFile(stream)
                {
                    FileName = ReceivedFile.Name
                };
                var send = await Bot.SendDocumentAsync(e.Message.Chat.Id, iof, "File");
            }
        }

        public static async void InsertTwo(MessageEventArgs e)
        {
            var doc = await Bot.GetFileAsync(e.Message.Document.FileId);
            var name = $@"C:\Users\Кристина\Desktop\{e.Message.Document.FileName}";
            FileStream fs = new FileStream(name, FileMode.Create);
            await Bot.GetInfoAndDownloadFileAsync(doc.FileId, fs);

            ReceivedFile = new FileInfo(name);
            ReceivedPackage = new ExcelPackage(ReceivedFile);

            var path = @$"C:\Users\Кристина\Desktop\ResultWithInsertion{DateTime.Now.ToShortTimeString()}.xlsx";
            NewFile = new FileInfo(path);
            NewPackage = new ExcelPackage(NewFile);
            for (int k = 0; k < ReceivedPackage.Workbook.Worksheets.Count; k++)
            {
                ReceivedWorksheet = ReceivedPackage.Workbook.Worksheets[k];
                Columns = ReceivedWorksheet.Dimension?.Columns ?? 0;
                NewWorksheet = NewPackage.Workbook.Worksheets.Add(ReceivedWorksheet.Name);
                await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);

                Rows = ReceivedWorksheet.Dimension?.Rows ?? 0;
                Console.WriteLine($"Sheet {ReceivedWorksheet.Name}");
                for (int j = 1; j <= Columns; j++)
                {
                    Console.WriteLine($"Column {j}");
                    if (ReceivedWorksheet.Cells[1, j].Value?.ToString() == "Цена" ||
                        ReceivedWorksheet.Cells[1, j].Value?.ToString() == "Цена(ю.)")
                    {
                        for (int i = 1; i <= Rows; i++)
                        {
                            ReceivedWorksheet.Cells[i, Columns].Value = 
                                (((IConvertible)ReceivedWorksheet.Cells[i, j].Value).ToDouble(null) * 6.3 / 1.3).ToString();
                            await NewPackage.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);
                        }
                    }
                }

                using var stream1 = File.OpenRead(name);
                InputOnlineFile iof1 = new InputOnlineFile(stream1)
                {
                    FileName = ReceivedFile.Name
                };
                await Bot.SendDocumentAsync(e.Message.Chat.Id, iof1, "File");

            }
            using var stream = File.OpenRead(path);
            InputOnlineFile iof = new InputOnlineFile(stream)
            {
                FileName = NewFile.Name
            };
            await Bot.SendDocumentAsync(e.Message.Chat.Id, iof, "File");
        }
        class Commands
        {
            public static string DeleteAll = "Reply to this message, please :) \n"
                + "Put a excel-file for deleting all columns exept attributes and count 🌸🌸🌸";
            public static string DeleteAllExceptPrice = "Reply to this message, please :) \n"
                + "Put a excel-file for deleting columns exept price, count and attributes 🌸🌸🌸";
            public static string AddPurchase = "Reply to this message, please :) \n"
                + "Put a excel-file for adding purchase and percents🌸🌸🌸";
        }
    }
}
