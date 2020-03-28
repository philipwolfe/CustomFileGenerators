using System;
using System.IO;
using System.Runtime.InteropServices;
using CustomFileGenerators;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestProject1
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		[DeploymentItem("Input.txt")]
		[DeploymentItem("Output.txt")]
		public void TestMethod1()
		{
			IVsSingleFileGenerator gen = new ObservablePropertyGenerator();
			uint output;
			var mem = new IntPtr[1];

			gen.Generate(@"input.txt", File.ReadAllText(@"input.txt"), "", mem, out output, null);

			var contents = new byte[output]; 
			Marshal.Copy(mem[0], contents, 0, (int)output);

			//File.WriteAllBytes("run.txt", contents);

			var reader = new StreamReader(new MemoryStream(contents));
			var fileReader = new StreamReader("output.txt");

			string outputLine = reader.ReadLine();
			while (reader.Peek() != 0)
			{
				string controlLine = fileReader.ReadLine();

				if (outputLine == null && controlLine == null)
					break;

				if(!controlLine.Equals(outputLine))
					throw new Exception("'" + outputLine + "' did not match expected '" + controlLine + "'");

				outputLine = reader.ReadLine();
			}
		}
		
		[TestMethod]
		public void InlineCommentTest1()
		{
			var gen = new ObservablePropertyGenerator();
			Assert.AreEqual(-1, gen.LocationOfInlineComment(""));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("var h = 1;"));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("var h = @\"c:\\\";"));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("var h = @\"\\\\server\\share\";"));
			Assert.AreEqual(10, gen.LocationOfInlineComment("var h = 1; //"));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("var h = \"//\";"));
			Assert.AreEqual(14, gen.LocationOfInlineComment("if (something) // slkfj \" alsdkfj \" lskfj "));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("var u = \"http://lfkj.sldkfj\";"));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("var d = \"\\\"\" + \"//\" + \"\\\"\";"));
			Assert.AreEqual(26, gen.LocationOfInlineComment("var d = \"\\\"\" + \"t\" + \"\\\"\"; // used to be \"\\t\\"));
			Assert.AreEqual(-1, gen.LocationOfInlineComment("[GenerateObservableProperty] /* inline */ private Collateral _selectedCollateral;"));
		}
	}
}
