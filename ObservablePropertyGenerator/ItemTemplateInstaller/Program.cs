using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Permissions;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ItemTemplateInstaller
{
	[RunInstaller(true)]
	public class Program : Installer
	{
		private const string FILE_NAME = "ObservablePropertyItemTemplate.zip";
		private const string INSTALL_DIR = "install.dir";
		private const string SUB_FOLDER = @"ItemTemplates\CSharp\Code\1033\";
		private const string MESSAGEBOX_TITLE = "CustomFileGenerators Item Templates Installation";
		private const string CUSTOM_ACTION_HEADER = "Installer Custom Action: ";
		private const string INSTALL_SUCCESS = "CustomActionInstallSuccess";
		private const string LOG_FILE_NAME = "ItemTemplateInstaller.log";

		[SecurityPermission(SecurityAction.Demand)]
		public override void Install(IDictionary savedState)
		{
			base.Install(savedState);

			Context = new InstallContext(Path.Combine(TargetDir, LOG_FILE_NAME), new string[]{});

			Context.LogMessage(CUSTOM_ACTION_HEADER + "Calling InstallInternal.");

			InstallInternal(savedState);
		}

		[SecurityPermission(SecurityAction.Demand)]
		public override void Rollback(IDictionary savedState)
		{
			base.Rollback(savedState);

			UninstallInternal(savedState);
		}

		[SecurityPermission(SecurityAction.Demand)]
		public override void Uninstall(IDictionary savedState)
		{
			base.Uninstall(savedState);

			UninstallInternal(savedState);
		}

		private void InstallInternal(IDictionary savedState)
		{
			Context.LogMessage(CUSTOM_ACTION_HEADER + "In InstallInternal.");

			Context.LogMessage(CUSTOM_ACTION_HEADER + "Locating Visual Studio 2010 install location...");

			object installDir = RegistryLocation();

			if (installDir == null)
			{
				Context.LogMessage(CUSTOM_ACTION_HEADER + "Automatic locating failed...");

				MessageBox.Show("The installer was unable to automatically locate your Visual Studio 2010 installation.  Please locate the msenv.dll file manually.", MESSAGEBOX_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
			
				OpenFileDialog dialog = new OpenFileDialog();
				dialog.Title = "Please locate the msenv.dll file manually.";
				dialog.Filter = "msenv.dll|msenv.dll";

				DialogResult result;
				bool retry = true;
				do
				{
					result = dialog.ShowDialog();

					if (result == DialogResult.OK)
					{
						if (File.Exists(dialog.FileName))
						{
							installDir = Path.GetDirectoryName(dialog.FileName);
							savedState.Add(INSTALL_DIR, installDir.ToString());
							retry = false;
						}
					}
					else
					{
						retry = false;
					}
				} while (retry);
			}
			else
			{
				Context.LogMessage(CUSTOM_ACTION_HEADER + "Visual Studio 2010 location is at " + installDir);
			}

			if (installDir != null)
			{
				Context.LogMessage(CUSTOM_ACTION_HEADER + "Copying templates...");
				File.Copy(Path.Combine(TargetDir, FILE_NAME), Path.Combine(installDir.ToString(), SUB_FOLDER, FILE_NAME), true);

				Context.LogMessage(CUSTOM_ACTION_HEADER + "Registering templates with Visual Studio 2010...");
				RegistrationProcess();

				Context.LogMessage(CUSTOM_ACTION_HEADER + "Registration complete...");
			}
			else
			{
				Context.LogMessage("Visual Studio 2010 installation directory not found.  Templates must be installed manually.");
				MessageBox.Show("The CustomFileGenerator item templates were not installed into Visual Studio 2010.  Perform an internet search on 'InstallVSTemplates' to install them manually.", MESSAGEBOX_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
			}

			savedState.Add(INSTALL_SUCCESS, true);
		}

		private void UninstallInternal(IDictionary savedState)
		{
			if (!savedState.Contains(INSTALL_SUCCESS))
				return;

			object installDir = RegistryLocation();

			if(installDir == null)
			{
				if(savedState.Contains(INSTALL_DIR))
				{
					installDir = savedState[INSTALL_DIR];
				}
			}

			if(installDir != null)
			{
				File.Delete(Path.Combine(installDir.ToString(), SUB_FOLDER, FILE_NAME));
				RegistrationProcess();
			}
		}

		private object RegistryLocation()
		{
			if (Environment.Is64BitOperatingSystem)
			{
				return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\10.0", "InstallDir", null);
			}
			else
			{
				return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\10.0", "InstallDir", null);
			}
		}

		private void RegistrationProcess()
		{
			var process = Process.Start("devenv.exe", "/installvstemplates");

			if(process != null)
				process.WaitForExit();
		}

		private string TargetDir
		{
			get
			{
				var targetDir = Context.Parameters["TargetDir"];
				if (targetDir == null)
				{
					var location = Assembly.GetExecutingAssembly().Location;
					return Path.GetDirectoryName(location);
				}
				else
				{
					return targetDir;
				}
			}
		}
	}
}