using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using QRCoder;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;

namespace WindowsFormsBLEserver
{
    class Program
    {
        static byte cnt = 0;
        static readonly string ConfigFilePath = "config.json";

        [STAThread]
        static void Main(string[] args)
        {
            //// 3秒待機
            //Thread.Sleep(3000);

            //Console.WriteLine("\\left(x+a\\right)^n=\\sum{k=0}^{n}{\\binom{n}{k}x^ka^{n-k}}");
            var (uuid1, uuid2) = LoadOrCreateUUID();


            //Console.WriteLine("アプリUUID: " + appUUID);

            // 非同期 BLE 処理の実行
            Task.Run(AsyncMain);

            // フォームを表示（QRコードを含む）
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new QRCodeForm(uuid1,uuid2));
        }

        static (Guid, Guid) LoadOrCreateUUID()
        {
            if (File.Exists(ConfigFilePath))
            {
                // ファイルからUUIDのペアを読み込む
                string json = File.ReadAllText(ConfigFilePath);
                Guid[] guids = JsonConvert.DeserializeObject<Guid[]>(json);

                if (guids.Length == 2)
                {
                    return (guids[0], guids[1]);
                }
                else
                {
                    throw new InvalidOperationException("UUIDペアが無効です");
                }
            }
            else
            {
                // 新しいUUIDを生成
                Guid newUuid1 = Guid.NewGuid();
                Guid newUuid2 = Guid.NewGuid();

                // 生成したUUIDペアを保存
                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(new Guid[] { newUuid1, newUuid2 }));

                // 新しいUUIDペアを返す
                return (newUuid1, newUuid2);
            }
        }


        static async Task AsyncMain()
        {
            // UUIDペアを読み込む（ファイルがない場合は生成）
            var (serviceUuid, characteristicUuid) = LoadOrCreateUUID();

            // GATT サービスの作成
            var gattServiceProviderResult = await GattServiceProvider.CreateAsync(serviceUuid);
            if (gattServiceProviderResult.Error != BluetoothError.Success)
            {
                Console.WriteLine("GATT Serviceの起動に失敗(Bluetooth LE対応デバイスがない?)");
                return;
            }

            var gattServiceProvider = gattServiceProviderResult.ServiceProvider;

            // GATT 特徴のパラメータ設定
            var cReadWriteParam = new GattLocalCharacteristicParameters
            {
                CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Write | GattCharacteristicProperties.Notify,
                ReadProtectionLevel = GattProtectionLevel.Plain,
                WriteProtectionLevel = GattProtectionLevel.Plain,
                UserDescription = "cReadWrite"
            };

            var cReadWrite = await gattServiceProvider.Service.CreateCharacteristicAsync(characteristicUuid, cReadWriteParam);

            cReadWrite.Characteristic.WriteRequested += async (GattLocalCharacteristic sender, GattWriteRequestedEventArgs args) =>
            {
                var deferral = args.GetDeferral();
                var request = await args.GetRequestAsync();
                var buffer = request.Value.ToArray();
                string receivedText = Encoding.UTF8.GetString(buffer);

                Console.WriteLine("受信: " + receivedText);
                if (receivedText == "★")
                {
                    string clipboardText = await CopySendText();
                    //Thread.Sleep(100); // コピーが完了するのを待つ
                    if (!string.IsNullOrEmpty(clipboardText))
                    {
                        byte[] utf8Data = Encoding.UTF8.GetBytes(clipboardText);
                        await sender.NotifyValueAsync(utf8Data.AsBuffer()); // BLE で送信
                        Console.WriteLine("BLE送信: " + clipboardText);
                    }
                }
                else {
                    SetClipboardText(receivedText);
                    PasteClipboardText();
                }


                if (request.Option == GattWriteOption.WriteWithResponse)
                {
                    request.Respond();
                }
                deferral.Complete();
            };

            var gattServiceProviderAdvertisingParameters = new GattServiceProviderAdvertisingParameters
            {
                IsConnectable = true,
                IsDiscoverable = true
            };

            gattServiceProvider.StartAdvertising(gattServiceProviderAdvertisingParameters);
            Console.WriteLine("StartAdvertising...");

            while (true)
            {
                cnt++;
            }
        }

        static void SetClipboardText(string text)
        {
            Thread staThread = new Thread(() =>
            {
                try
                {
                    Clipboard.SetText(text, TextDataFormat.UnicodeText);
                    Console.WriteLine("クリップボードにコピー: " + text);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("クリップボードエラー: " + ex.Message);
                }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
        }

        static void PasteClipboardText()
        {
            Thread staThread = new Thread(() =>
            {
                try
                {
                    //ペースト^V以外のコマンドは条件分岐の必要あり　もしくはSwift側から文字列内部に入れるか

                    SendKeys.SendWait("%n");
                   // Thread.Sleep(10);
                    SendKeys.SendWait("e");
                   // Thread.Sleep(10);
                    SendKeys.SendWait("i");
                   // Thread.Sleep(10);
                    SendKeys.SendWait("^v");
                    SendKeys.SendWait("{ENTER}");
                    Console.WriteLine("ペースト実行");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ペーストエラー: " + ex.Message);
                }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
        }
        //static string CopySendText()
        //{

        //    string clipboardText = string.Empty;
        //    Thread staThread = new Thread(() =>
        //    {
        //        try
        //        {
        //            //ペースト^V以外のコマンドは条件分岐の必要あり　もしくはSwift側から文字列内部に入れるか
        //            SendKeys.SendWait("^c");
        //            Thread.Sleep(100); // コピーが完了するのを待つ
        //            clipboardText = Clipboard.GetText(TextDataFormat.UnicodeText); // クリップボードから取得
        //            Console.WriteLine("コピーしたテキスト: " + clipboardText);
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine("ペーストエラー: " + ex.Message);
        //        }
        //    });

        //    staThread.SetApartmentState(ApartmentState.STA);
        //    staThread.Start();
        //    staThread.Join();

        //    return clipboardText;
        //}

        static async Task<string> CopySendText()
        {
            var tcs = new TaskCompletionSource<string>();

            Thread staThread = new Thread(() =>
            {
                try
                {
                    SendKeys.SendWait("^c"); // Ctrl+C を送信
                   // Thread.Sleep(1000); // コピーの完了を待つ

                    string clipboardText = Clipboard.GetText(TextDataFormat.UnicodeText); // クリップボードから取得
                    Console.WriteLine("コピーしたテキスト: " + clipboardText);

                    tcs.SetResult(clipboardText);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("コピーエラー: " + ex.Message);
                    tcs.SetResult(string.Empty); // エラー時は空文字を返す
                }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();

            return await tcs.Task;
        }


    }

    /// <summary>
    /// UUID を QR コードとして表示するフォーム
    /// </summary>
    public class QRCodeForm : Form
    {
        private PictureBox pictureBox;
        private Label label;
        private Guid UUID1;
        private Guid UUID2;
        private string combinedUUID;

        public QRCodeForm(Guid uuid1, Guid uuid2)
        {
            this.UUID1 = uuid1;
            this.UUID2 = uuid2;
            this.combinedUUID = uuid1.ToString() + "/" + uuid2.ToString();  // UUIDを結合

            this.Text = "UUID QRコード表示";
            this.Width = 400;
            this.Height = 450;

            label = new Label
            {
                Text = "UUID: " + combinedUUID.ToString(),
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter
            };

            pictureBox = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.Zoom,
                Dock = DockStyle.Fill
            };

            this.Controls.Add(pictureBox);
            this.Controls.Add(label);
            DisplayQRCode();
        }

        private void DisplayQRCode()
        {
            using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
            {
                using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(combinedUUID.ToString(), QRCodeGenerator.ECCLevel.Q))
                {
                    using (QRCode qrCode = new QRCode(qrCodeData))
                    {
                        pictureBox.Image = qrCode.GetGraphic(10);
                    }
                }
            }
        }
    }
}
