using System.Collections;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Permissions;
using System.Text;
using System.Windows.Forms;

namespace ItemTemplateInstaller
{
	public abstract class VSTemplateInstallerBase : Installer
	{
		private const string FILE_NAME = "ObservablePropertyItemTemplate.zip";
		private const string INSTALL_DIR = "install.dir";
		private const string SUB_FOLDER = @"ItemTemplates\CSharp\Code\1033\";
		private const string MESSAGEBOX_TITLE = "CustomFileGenerators Item Templates Installation";
		private const string INSTALL_SUCCESS = "CustomActionInstallSuccess";

		//public VSTemplateInstallerBase(string logFileName) : base(logFileName)
		//{
		//}

		protected void InstallInternal(IDictionary stateSaver)
		{
			Context.LogMessage(CustomActionHeader + "In InstallInternal.");

			Context.LogMessage(CustomActionHeader + "Locating " + ProductTitle + " install location...");

			var installDir = InstallDir;

			if (!string.IsNullOrEmpty(installDir) && Directory.Exists(installDir))
			{
				var msenv_dll = Path.Combine(installDir, "msenv.dll");

				var file = FileVersionInfo.GetVersionInfo(msenv_dll);

				if (file.ProductMajorPart == VSVersionMajor)
				{
					stateSaver.Add(INSTALL_DIR, installDir);

					Context.LogMessage(CustomActionHeader + ProductTitle + " location is at " + installDir);

					CopyTemplates(installDir);

					ResetTemplates();

					stateSaver.Add(INSTALL_SUCCESS, true);
				}
				else
				{
					Context.LogMessage(CustomActionHeader + ProductTitle + " version mismatch.  Version found: ");
				}
			}
			else
			{
				Context.LogMessage(CustomActionHeader + ProductTitle + " installation directory not found.  Templates must be installed manually.");
			}
		}

		protected void UninstallInternal(IDictionary stateSaver)
		{
			if (stateSaver == null)
				return;

			if (!stateSaver.Contains(INSTALL_SUCCESS))
				return;

			var installDir = InstallDir;

			if (string.IsNullOrEmpty(installDir))
			{
				if (stateSaver.Contains(INSTALL_DIR))
				{
					installDir = stateSaver[INSTALL_DIR].ToString();
				}
			}

			if (!string.IsNullOrEmpty(installDir))
			{
				DeleteTemplates(installDir);
				ResetTemplates();
			}
		}

		protected abstract string InstallDir { get; }
		protected abstract string ProductTitle { get; }
		protected abstract string CustomActionHeader { get; }
		protected abstract int VSVersionMajor { get; }
		protected abstract string DevEnvLocation { get; }

		private void ResetTemplates()
		{
			Context.LogMessage(CustomActionHeader + "Resetting templates on " + ProductTitle);

			var process = Process.Start(DevEnvLocation, "/installvstemplates");

			if (process != null)
				process.WaitForExit();

			Context.LogMessage(CustomActionHeader + "Resetting templates complete.");
		}

		private void CopyTemplates(string installDir)
		{
			Context.LogMessage(CustomActionHeader + "Copying templates");
			File.Copy(Path.Combine(TargetDir, FILE_NAME), Path.Combine(installDir, SUB_FOLDER, FILE_NAME), true);
		}

		private void DeleteTemplates(string installDir)
		{
			Context.LogMessage(CustomActionHeader + "Deleting templates");
			File.Delete(Path.Combine(installDir, SUB_FOLDER, FILE_NAME));
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
