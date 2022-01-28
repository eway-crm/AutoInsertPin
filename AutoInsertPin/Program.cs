using Microsoft.Test.Input;
using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace AutoInsertPin
{
    class Program
    {
        static async Task Main(string[] args)
        {
            SecureString securePin = new SecureString();
            if (args.Length == 0)
            {
                Console.WriteLine("Type PIN (Enter to end typing): ");
                string pin = string.Empty;
                while (true)
                {
                    ConsoleKeyInfo keyInfo = Console.ReadKey(true);
                    if (keyInfo.Key == ConsoleKey.Enter)
                        break;

                    securePin.AppendChar(keyInfo.KeyChar);
                }
            }
            else
            {
                Array.ForEach(args[0].ToCharArray(), securePin.AppendChar);
            }

            securePin.MakeReadOnly();

            Console.WriteLine("Waiting for windows. Press Ctrl+C to finish monitoring windows...");

            while (true)
            {
                try
                {
                    var window = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Windows Security"));
                    if (window == null)
                    {
                        await Task.Delay(1000);
                        continue;

                    }
                    await Task.Delay(1000);

                    Console.WriteLine("Located window");

                    var controls = window.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "PIN"));
                    if (controls == null || controls.Count != 3)
                        continue;

                    controls[2].SetFocus();
                    SendKeys.SendWait(ConvertToPlainString(securePin));
                    await Task.Delay(250);

                    var okControl = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "OK"));
                    if (okControl == null)
                        continue;

                    Mouse.MoveTo(new System.Drawing.Point((int)okControl.GetClickablePoint().X, (int)okControl.GetClickablePoint().Y));
                    Mouse.Click(MouseButton.Left);

                    await Task.Delay(10000);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message);
                }
            }
        }

        public static string ConvertToPlainString(SecureString input)
        {
            IntPtr stringPointer = Marshal.SecureStringToBSTR(input);
            string plain = Marshal.PtrToStringBSTR(stringPointer);
            Marshal.ZeroFreeBSTR(stringPointer);

            return plain;
        }
    }
}
