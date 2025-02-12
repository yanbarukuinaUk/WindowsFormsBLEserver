using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;

namespace WindowsFormsBLEserver
{
    class Program
    {
        static byte cnt = 0;

        [STAThread]
        static void Main(string[] args)
        {
            // 3秒待機
            Thread.Sleep(3000);

            Console.WriteLine("\\left(x+a\\right)^n=\\sum{k=0}^{n}{\\binom{n}{k}x^ka^{n-k}}");

            // Async-awaitを使えるようにTask内で実行
            Task.Run(AsyncMain).Wait();
        }

        static async Task AsyncMain()
        {
            var gattServiceProviderResult = await GattServiceProvider.CreateAsync(new Guid("00000000-0000-4000-A000-000000000000"));
            if (gattServiceProviderResult.Error != BluetoothError.Success)
            {
                Console.WriteLine("GATT Serviceの起動に失敗(Bluetooth LE対応デバイスがない?)");
                return;
            }

            var gattServiceProvider = gattServiceProviderResult.ServiceProvider;

            var cReadWriteParam = new GattLocalCharacteristicParameters
            {
                CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Write | GattCharacteristicProperties.Notify,
                ReadProtectionLevel = GattProtectionLevel.Plain,
                WriteProtectionLevel = GattProtectionLevel.Plain,
                UserDescription = "cReadWrite"
            };

            var cReadWrite = await gattServiceProvider.Service.CreateCharacteristicAsync(new Guid("00000000-0000-4000-A000-000000000001"), cReadWriteParam);

            // クライアントからのデータ受信処理 (UTF-8 でデコード)
            cReadWrite.Characteristic.WriteRequested += async (GattLocalCharacteristic sender, GattWriteRequestedEventArgs args) =>
            {
                var deferral = args.GetDeferral();
                var request = await args.GetRequestAsync();
                var buffer = request.Value.ToArray();
                string receivedText = Encoding.UTF8.GetString(buffer); // UTF-8 デコード

                Console.WriteLine("受信: " + receivedText);
                SetClipboardText(receivedText);
                PasteClipboardText();

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

            // 定期的にクリップボードへコピー＆ペースト
            while (true)
            {
              //  string message = "通知: " + cnt.ToString();
              //  SetClipboardText(message);
              //  PasteClipboardText();

              //  await Task.Delay(1000);
                cnt++;
            }
        }

        /// <summary>
        /// クリップボードにテキストをコピー (UTF-8 を考慮, STA スレッドで実行)
        /// </summary>
        static void SetClipboardText(string text)
        {
            Thread staThread = new Thread(() =>
            {
                try
                {
                    Clipboard.SetText(text, TextDataFormat.UnicodeText); // UnicodeText で UTF-8 を考慮
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

        /// <summary>
        /// クリップボードの内容を貼り付け (STA スレッドで実行)
        /// </summary>
        static void PasteClipboardText()
        {
            Thread staThread = new Thread(() =>
            {
                try
                {
                    //ペースト^V以外のコマンドは条件分岐の必要あり　もしくはSwift側から文字列内部に入れるか
                    SendKeys.SendWait("+%-");
                    SendKeys.SendWait("^v"); // Ctrl + V で貼り付け
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
    }
}