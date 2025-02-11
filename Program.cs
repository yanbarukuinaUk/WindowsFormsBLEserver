
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Bluetooth;
using System.IO;
using System.Threading;

namespace WindowsFormsBLEserver
{

    class Program
    {
        static byte cnt = 0;
        static void Main(string[] args)
        {
            Console.WriteLine("IMEをオンにします...");
            SetIME(true);  // IMEをオンにする

            Console.WriteLine("3秒後に'😀'を入力します...");
            System.Threading.Thread.Sleep(3000); // 3秒待つ

            SendKeys.SendWait("😀vghjghfghfgygyhghgjhghjghj");  // 絵文字を入力
            Console.WriteLine("入力完了！");

            //Async-awaitを使えるようにTask内で実行する
            Task.Run(AsyncMain).Wait();
        }

        //IMEをオン/オフする
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("imm32.dll")]
        static extern IntPtr ImmGetContext(IntPtr hWnd);

        [DllImport("imm32.dll")]
        static extern bool ImmSetOpenStatus(IntPtr hIMC, bool bOpen);

        static void SetIME(bool enable)
        {
            IntPtr hwnd = GetForegroundWindow();
            IntPtr hIMC = ImmGetContext(hwnd);
            ImmSetOpenStatus(hIMC, enable);
        }


        static async Task AsyncMain()
        {
            //GattServiceProviderを指定のUUIDで初期化する
            //システムに独自のBLE GATTサービスを追加する。
            //失敗する場合、たいていBluetooth LE対応の環境ではない。(デスクトップPC、古いPC、Bluetooth 3以前のドングルを使用中など)
            //または、システムにより禁止された予約済みUUIDを使用している。

            var gattServiceProviderResult = await GattServiceProvider.CreateAsync(new Guid("00000000-0000-4000-A000-000000000000"));
            if (gattServiceProviderResult.Error != BluetoothError.Success)
            {
                Console.WriteLine("GATT Serviceの起動に失敗(Bluetooth LE対応デバイスがない?)");
                return;
            }

            var gattServiceProvider = gattServiceProviderResult.ServiceProvider;

            //---

            //ローカルキャラクタリスティック(外部から読み書き可能な値)を定義する
            var cReadWriteParam = new GattLocalCharacteristicParameters
            {
                CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Write | GattCharacteristicProperties.Notify, //読み込み & 書き込み & 通知購読可能
                ReadProtectionLevel = GattProtectionLevel.Plain, //誰でも読み込み可能
                WriteProtectionLevel = GattProtectionLevel.Plain, //誰でも書き込み可能
                UserDescription = "cReadWrite" //ユーザーに見える説明(BLEツールを使って読むことができる)
            };

            //定義した情報をもとに、指定のUUIDでサービスに登録する
            var cReadWrite = await gattServiceProvider.Service.CreateCharacteristicAsync(new Guid("00000000-0000-4000-A000-000000000001"), cReadWriteParam);

            //読み込みが発生したときのコールバック定義
            cReadWrite.Characteristic.ReadRequested += async (GattLocalCharacteristic sender, GattReadRequestedEventArgs args) =>
            {
                //接続中のデバイスから読み込まれた
                Console.WriteLine("Read request from " + args.Session.DeviceId.Id);
                var deferral = args.GetDeferral(); //非同期処理完了を知らせるためのDeferral (awaitを使うため)

                var request = await args.GetRequestAsync(); //リクエストを取得

                byte[] buf = new byte[1] { cnt };  //返却値を準備(Streamでもいいが、単純のためにbyte[]を使用)
                request.RespondWithValue(buf.AsBuffer()); //返却

                deferral.Complete(); //非同期完了を通知
            };

            //書き込みが発生したときのコールバック定義
            cReadWrite.Characteristic.WriteRequested += async (GattLocalCharacteristic sender, GattWriteRequestedEventArgs args) =>
            {
                //接続中のデバイスから書き込まれた
                Console.WriteLine("Write request from " + args.Session.DeviceId.Id);
                var deferral = args.GetDeferral(); //非同期処理完了を知らせるためのDeferral (awaitを使うため)

                var request = await args.GetRequestAsync(); //リクエストを取得

                var stream = request.Value.AsStream(); //streamを取得

                //1byteずつ読み込んで表示
                int d = 0;
                while ((d = stream.ReadByte()) != -1)
                {
                    cnt = (byte)d;
                    Console.Write(d.ToString("X"));
                    Console.Write(",");
                    SendKeys.SendWait(d.ToString("X"));
                }
                Console.WriteLine();

                if (request.Option == GattWriteOption.WriteWithResponse)
                {
                    request.Respond(); //送信側が応答欲しい場合は応答を返す(これをしないと送信側がエラーになる)
                    //System.Exception: 要求された属性要求で思いもよらないエラーが発生したため、要求されたとおりに完了することができませんでした。 (HRESULT からの例外:0x8065000E) の原因になる
                }

                deferral.Complete(); //非同期完了を通知
            };

            //購読者の増減が発生したときのコールバック定義
            cReadWrite.Characteristic.SubscribedClientsChanged += async (GattLocalCharacteristic sender, object args) =>
            {
                //購読者が増えた/減った
                Console.WriteLine("Subscribe Changed(Notify)");
                foreach (var c in sender.SubscribedClients)
                {
                    Console.WriteLine("- Device: " + c.Session.DeviceId.Id);
                }
                Console.WriteLine("- DeviceEnd");
            };

            //サービスをアドバタイジングするパラメータ
            var gattServiceProviderAdvertisingParameters = new GattServiceProviderAdvertisingParameters
            {
                IsConnectable = true, //接続可能
                IsDiscoverable = true //検出可能
            };

            //アドバタイジング開始
            gattServiceProvider.StartAdvertising(gattServiceProviderAdvertisingParameters);

            Console.WriteLine("StartAdvertising...");

            //1秒おきにカウントアップ値をNotifyする

            while (true)
            {

                byte[] bufN = new byte[1] { cnt };
                await cReadWrite.Characteristic.NotifyValueAsync(bufN.AsBuffer());
                Console.WriteLine("Notify " + cnt.ToString("X"));

                await Task.Delay(1000);

                cnt++;
            }

        }

    }
    
}

