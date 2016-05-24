using System;
using System.Runtime.InteropServices;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using System.Threading;
using System.IO;

namespace Win_Keyboard_Layout_Remover
{
	internal class Program
	{
		/// <summary>
		/// Logs to file and prints to console a <c>string</c>.
		/// </summary>
		/// <param name="str">The string to write log and print to the console.</param>
		/// <seealso cref="File.AppendAllText(string, string)"/><br/>
		/// <seealso cref="Console.WriteLine(string)"/>
		public static void Logger(object str)
		{
			str = str.ToString();
			string d = DateTime.Now.ToString(CultureInfo.InvariantCulture);
			File.AppendAllText("log.log", $"[{d}]\t{str}\n");
			Console.WriteLine($"[{d}] {str}");
		}

		/// <summary>
		/// Detects a change in the keyboard layout list, removes the new addition, and sets the active language to the one prior to the change.
		/// </summary>
		/// <seealso cref="PowerShellExecutor"/><br/>
		/// <seealso cref="PowerShellExecutor.ExecuteSynchronously"/><br/>
		/// <seealso cref="System.Collections.Generic.List&lt;T/&gt;"/><br/>
		/// <seealso cref="GetSet.GetCurrentKeyboardLayout()"/><br/>
		/// <seealso cref="GetSet.SetInputMethod(int)"/><br/>
		private static void Main()
		{
			PowerShellExecutor ps = new PowerShellExecutor();
			dynamic list = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;

			while (true)
			{
				dynamic newList = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;
				dynamic replaceList = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;
				int layoutId = GetSet.GetCurrentKeyboardLayout();

				if (list.Count != newList.Count)
				{
					Logger($"New language was detected while window \"{GetSet.GetActiveWindowTitle()}\" was active.");
					Boolean found = false;
					for (int i = 0; i < newList.Count; i++)
					{
						var newItem = newList[i];
						foreach (var oldItem in list) if (newItem.LanguageTag == oldItem.LanguageTag) found = true;
						if (!found) replaceList.RemoveAt(i);
					}
					ps.ExecuteSynchronously("param($finalList) Set-WinUserLanguageList($finalList) -Force", true, "finalList", replaceList);
					Logger("New language removed.");
					GetSet.SetInputMethod(layoutId);
				}
				list = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;
				Thread.Sleep(1000);
			}
		}
	}

	internal class GetSet
	{

		[DllImport("user32.dll")]
		private static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll")]
		private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

		[DllImport("user32.dll")]
		private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr process);

		[DllImport("user32.dll")]
		private static extern IntPtr GetKeyboardLayout(uint thread);

		[DllImport("user32.dll")]
		private static extern IntPtr PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll")]
		private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint flags);

		/// <summary>
		/// Gets the culture identifier for the active keyboard layout for the current foreground window.
		/// </summary>
		/// <returns>the culture identifier for the active keyboard layout for the current foreground window.</returns>
		/// <seealso cref="CultureInfo.LCID"></seealso>
		public static int GetCurrentKeyboardLayout()
		{
			try
			{
				IntPtr foregroundWindow = GetForegroundWindow();
				uint foregroundProcess = GetWindowThreadProcessId(foregroundWindow, IntPtr.Zero);
				int keyboardLayout = GetKeyboardLayout(foregroundProcess).ToInt32() & 0xFFFF;
				return keyboardLayout;
			}
			catch (Exception)
			{
				Program.Logger($"Invalid locale ID. Defaulting to {GetCultureInfo(1053).DisplayName}.");
				return 1053;
			}
		}

		/// <summary>
		/// Gets a <c>CultureInfo</c> object based on a culture identifier.
		/// </summary>
		/// <param name="layoutId">The culture identifier.</param>
		/// <returns></returns>
		/// <seealso cref="CultureInfo.LCID"></seealso>
		public static CultureInfo GetCultureInfo(int layoutId)
		{
			return new CultureInfo(layoutId);
		}

		public static string GetActiveWindowTitle()
		{
			const int c = 256;
			StringBuilder buff = new StringBuilder(c);
			IntPtr handle = GetForegroundWindow();

			return GetWindowText(handle, buff, c) > 0 ? buff.ToString() : null;
		}

		/// <summary>
		/// Sets the active keyboard layout based on a culture identifier.
		/// </summary>
		/// <param name="layoutId">The The culture identifier.</param>
		/// <seealso cref="CultureInfo.LCID"/>
		public static void SetInputMethod(int layoutId)
		{
			switch (layoutId)
			{
				case 1053:
					PostMessage(GetForegroundWindow(), 0x0050, (IntPtr)0, LoadKeyboardLayout("A000041D", 1));
					Program.Logger($"Active language set to {GetCultureInfo(layoutId).DisplayName}.");
					break;
				case 1049:
					PostMessage(GetForegroundWindow(), 0x0050, (IntPtr)0, LoadKeyboardLayout("A0000419", 1));
					Program.Logger($"Active language set to {GetCultureInfo(layoutId).DisplayName}.");
					break;
				default:
					PostMessage(GetForegroundWindow(), 0x0050, (IntPtr)0, LoadKeyboardLayout("A000041D", 1));
					Program.Logger($"Error setting the active language. Defaulting to {GetCultureInfo(1053).DisplayName}.");
					break;
			}
		}
	}

	internal class PowerShellExecutor
	{
		/// <summary>
		/// Synchronously invokes a PowerShell script and returns the object that the PowerShell script outputs.
		/// </summary>
		/// <param name="script">The PowerShell script to invoke.</param>
		/// <param name="b">If true <c>true</c>, adds value <paramref name="pVal"/> to parameter <paramref name="p"/>.</param>
		/// <param name="p">The name of the parameter for which to add value <paramref name="pVal"/>.</param>
		/// <param name="pVal">The value to add to parameter <paramref name="p"/>.</param>
		/// <returns><c>PSObject</c> outputted by the PowerShell <paramref name="script"/>.</returns>
		/// <seealso cref="PowerShell"/><br/>
		/// <seealso cref="PSObject"/>
		public PSObject ExecuteSynchronously(string script, Boolean b, string p, object pVal)
		{
			using (PowerShell PowerShellInstance = PowerShell.Create())
			{
				PowerShellInstance.AddScript(script);
				if (b) PowerShellInstance.AddParameter(p, pVal);
				Collection<PSObject> PSOutput = PowerShellInstance.Invoke();
				foreach (PSObject outputItem in PSOutput) if (outputItem != null) return outputItem;
			}
			return null;
		}
	}
}
