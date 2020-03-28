using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TextTemplating.VSHost;

namespace CustomFileGenerators
{
	[Guid(Constants.CLASS_GUID)]
	[ComVisible(true)]
	public class ObservablePropertyGenerator : BaseCodeGenerator
	{
		private int _indent;

		public override string GetDefaultExtension()
		{
			return Constants.DEFAULT_FILE_EXTENSION;
		}

		protected override byte[] GenerateCode(string inputFileName, string inputFileContent)
		{
			var output = GenerateCode(inputFileContent);
			var memoryStream = new MemoryStream();
			var writer = new StreamWriter(memoryStream, Encoding.Unicode);
			writer.Write(output);
			writer.Flush();
			return memoryStream.ToArray();
		}

		private string GenerateCode(string inputFileContent)
		{
			#region Remove single line comments
			var reader = new StringReader(inputFileContent);
			var parsedInput = new StringBuilder();

			string parsedLine;
			while ((parsedLine = reader.ReadLine()) != null)
			{
				parsedLine = parsedLine.Trim();

				// ignore lines that start with //
				if (!parsedLine.StartsWith("//"))
				{
					// look for an inline //
					int idxOfComment = LocationOfInlineComment(parsedLine);

					if (idxOfComment != -1)
						parsedLine = parsedLine.Substring(0, idxOfComment + 1); // read up to the //

					// add the new line
					parsedInput.AppendLine(parsedLine);
				}
			}

			// reset the inputFileContent
			inputFileContent = parsedInput.ToString();
			parsedInput.Clear();
			#endregion

			#region Remove Multiline comments

			// run the multi-line comment regex
			inputFileContent = Regex.Replace(inputFileContent, @"(\/\*[\w\'\s\r\n\*]*\*\/)|(\/\/[\w\s\']*)", string.Empty);

			#endregion

			_indent = 0;
			StringBuilder sb = new StringBuilder();

			sb.Append(string.Format(Constants.FILE_HEADER, GetType().Assembly.GetName().Version, Environment.Version));

			try
			{
				//bring in usings of input file
				foreach (Match match in Regex.Matches(inputFileContent, Constants.USING_REGEX))
				{
					sb.Append(match.Value + Environment.NewLine);
				}

				sb.Append(Environment.NewLine);

				//bring in first namespace of input file
				var namespaceLine = Regex.Match(inputFileContent, Constants.NAMESPACE_REGEX).Value;

				sb.AppendLine(namespaceLine);
				sb.AppendLine("{");
				_indent++;

				//bring in first class of input file
				var classLine = Regex.Match(inputFileContent, Constants.CLASS_REGEX).Value;
				//remove stuff after a : or where
				var colonIndex = classLine.IndexOf(':');
				if (colonIndex >= 0)
				{
					classLine = classLine.Substring(0, colonIndex - 1);
				}
				var whereIndex = classLine.IndexOf(" where ");
				if (whereIndex >= 0)
				{
					classLine = classLine.Substring(0, whereIndex - 1);
				}
				var className = classLine.Substring("class ".Length).Trim();

				var angleIndex = className.IndexOf('<');
				if (angleIndex >= 0)
				{
					className = className.Substring(0, angleIndex).Trim();
				}

				sb.AppendLine(Indent + "partial " + classLine.Trim());
				sb.Append(Indent + "{");
				_indent++;

				var properties = new List<string>();
				var propertyTypes = new Dictionary<string, string>();

				//find first attribute
				var attribIndex = NextIndexOfAttribute(inputFileContent, 0);

				while (attribIndex != -1)
				{
					var index = attribIndex + Constants.GOP_ATTRIBUTE.Length;

					var optionsFound = false;
					var properName = string.Empty;
					while (!optionsFound)
					{
						switch (inputFileContent[index])
						{
							case ',': // new attribute
								optionsFound = true;
								index = inputFileContent.IndexOf(']', index) + 1;
								break;
							case ']': // end of this attribute
								optionsFound = true;
								index++;
								break;
							case '"': // start of quoted option
								index++;
								var indexOfEndQuote = inputFileContent.IndexOf('"', index);

								if (indexOfEndQuote != -1)
								{
									properName = inputFileContent.Substring(index, indexOfEndQuote - index);
									index = inputFileContent.IndexOf(']', index) + 1;
								}
								optionsFound = true;
								break;
							default:
								index++;
								break;
						}
					}

					bool propertyFound = false;
					bool eof = false;
					string property = string.Empty;
					while (!propertyFound)
					{
						switch (inputFileContent[index])
						{
							case '}':
							case '{':
								// quit processing this attribute
								eof = true; //do not write property
								propertyFound = true; //break out of loop
								break;
							case ' ': //space
							case '\t': //tab
							case '\r': //newline
							case '\n':
							case ',': //attibute separator
								// move to the next char 
								index++;
								break;
							case '[':
								// another attribute
								index = inputFileContent.IndexOf(']', index) + 1;
								break;
							default:
								int nextLineIdx = inputFileContent.IndexOf("\r\n", index);
								if (nextLineIdx > 0 && !(nextLineIdx > inputFileContent.Length))
								{
									int lineEnd = inputFileContent.IndexOf(';', index) + 1;
									if (lineEnd > 0)
									{
										string line = inputFileContent.Substring(index, lineEnd - index);

										index = nextLineIdx + 2;
										property = line;
										propertyFound = true;
									}
									else
									{
										index = nextLineIdx + 1;
									}
								}
								else
								{
									index = nextLineIdx + 1;
								}
								break;
						}
					}

					if (!eof)
					{
						//preprocess property
						property = property.Trim();
						//property = property.Substring(0, property.IndexOf(';') + 1); //removes trailing comments
						property = Regex.Replace(property, @"\s+", " "); //normalize spaces

						int indexOfEquals = property.LastIndexOf('=');
						int indexOfSemiColon = property.IndexOf(';');

						int indexOfEndOfName;
						if (indexOfEquals == -1)
							indexOfEndOfName = indexOfSemiColon;
						else
						{
							// if there is a space before the = then go back one
							if (property[indexOfEquals - 1] == ' ')
								indexOfEquals--;

							indexOfEndOfName = Math.Min(indexOfEquals, indexOfSemiColon);
						}

						int indexOfLastSpace = property.Substring(0, indexOfEndOfName).LastIndexOf(' ') + 1;

						var propertyName = property.Substring(indexOfLastSpace, indexOfEndOfName - indexOfLastSpace);
						if (string.IsNullOrEmpty(properName))
							properName = ProperName(propertyName);

						if (properName != null) //invalid name
						{
							properties.Add(properName);

							var returnType = property.Substring(0, indexOfLastSpace).Trim();
							var indexOfAngleBracket = property.IndexOf('<');

							var indexOfSpaceBeforeReturnType = 0;
							if (indexOfAngleBracket >= 0)
							{
								indexOfSpaceBeforeReturnType = returnType.Substring(0, indexOfAngleBracket).LastIndexOf(' ');
								if (indexOfSpaceBeforeReturnType == -1)
									indexOfSpaceBeforeReturnType = 0;
							}
							else
								indexOfSpaceBeforeReturnType = returnType.Substring(0, indexOfLastSpace - 1).LastIndexOf(' ') + 1;
							returnType = returnType.Substring(indexOfSpaceBeforeReturnType, indexOfLastSpace - indexOfSpaceBeforeReturnType - 1).Trim();
							propertyTypes.Add(properName, returnType);
							var formattedProperty = PropertyFormat(Indent, returnType, properName, propertyName, className);

							sb.AppendLine(formattedProperty);
						}
					}
					attribIndex = NextIndexOfAttribute(inputFileContent, index);
				}

				if (properties.Count == 0)
                    return string.Format(Constants.FILE_HEADER, GetType().Assembly.GetName().Version, Environment.Version);

				sb.AppendLine();
				sb.AppendLine(Indent + "#region Extensibility Method Definitions");
				foreach (var prop in properties)
				{
					sb.AppendLine(string.Format("{0}partial void OnBefore{1}ValueChanging({2} value, System.ComponentModel.CancelEventArgs e);", Indent, prop, propertyTypes[prop]));
					sb.AppendLine(string.Format("{0}partial void OnBefore{1}RaisePropertyChanged();", Indent, prop));
					sb.AppendLine(string.Format("{0}partial void OnAfter{1}RaisePropertyChanged();", Indent, prop));
				}
				sb.AppendLine(Indent + "#endregion");

				sb.AppendLine();
				sb.AppendLine(Indent + "public static partial class " + className + "Properties");
				sb.AppendLine(Indent + "{");
				_indent++;
				foreach (var prop in properties)
					sb.AppendLine(Indent + "public const string " + prop + " = \"" + prop + "\";");
				_indent--;
				sb.AppendLine(Indent + "}");
				_indent--;
				sb.AppendLine(Indent + "}"); //class
				_indent--;
				sb.AppendLine(Indent + "}"); //namespace
				return sb.ToString();
			}
			catch (Exception ex)
			{
				return "/* Unable to generate code.  Exception information:" + Environment.NewLine + ex + Environment.NewLine + "*/";
			}
		}

		private string Indent
		{
			get
			{
				return new string('\t', _indent);
			}
		}

		private string ProperName(string name)
		{
			if (name[0] == '_')
			{
				switch (name.Length)
				{
					case 1:
						return null;
					case 2:
						return name[1].ToString().ToUpper();
					default:
						return name[1].ToString().ToUpper() + name.Substring(2);
				}
			}

			return name[0].ToString().ToUpper() + name.Substring(1);
		}

		public int LocationOfInlineComment(string inputString)
		{
			int idxOfComment = -1;
			bool quoted = false, escaping = false, commented = false;
			for (int i = 0; i < inputString.Length; i++)
			{
				switch (inputString[i])
				{
					case '"':
						if (escaping)
							escaping = false;
						else
							quoted = !quoted;
						break;
					case '\\':
						if (quoted)
							escaping = true;
						break;
					case '/':
						if (commented)
							if (!quoted)
							{
								idxOfComment = i - 2;
								goto exit;
							}
							else
								commented = false;
						else
							commented = true;
						break;
					default:
						commented = false;
						escaping = false;
						break;
				}
			exit:
				;
			}
			return idxOfComment;

		}

		/// <summary>
		/// Locates the next occurrence of the attribute
		/// </summary>
		/// <param name="inputFileContent"></param>
		/// <param name="startIndex"></param>
		/// <returns></returns>
		private int NextIndexOfAttribute(string inputFileContent, int startIndex)
		{
			int idx = inputFileContent.IndexOf(Constants.GOP_ATTRIBUTE, startIndex);

			if (idx == -1) // attribute not found
				return idx;

			int indexOfPreviousWord = idx;

			//Note: we need to expect some reasonable formatting for this part.
			/*
			 * public class GenerateObservablePropertyAttribute : Attribute
			 * {
			 *     public GenerateObservablePropertyAttribute()
			 *     {
			 *     
			 *     }
			 *     public GenerateObservablePropertyAttribute(string name)
			 *     {
			 *     
			 *     }
			 * }
			 */

			//don't count the space before the attribute
			if (inputFileContent[indexOfPreviousWord - 1] == ' ')
				indexOfPreviousWord = indexOfPreviousWord - 2;

			//find the text preceeding the attribute
			while (inputFileContent[indexOfPreviousWord] != ' ' && inputFileContent[indexOfPreviousWord] != '\n')
			{
				indexOfPreviousWord--;
			}

			string previousWord = inputFileContent.Substring(indexOfPreviousWord, idx - indexOfPreviousWord).Trim();

			switch (previousWord)
			{
				case "class":
				case "public":
				case "private":
				case "internal":
				case "protected":
					return NextIndexOfAttribute(inputFileContent, idx + 1);
				default:
					return idx;

			}

		}

		private readonly string _propertyFormat = @"
{0}public {1} {2}
{0}{{
{0}	get
{0}	{{
{0}		return {3};
{0}	}}
{0}	set
{0}	{{
{0}		if (!Equals({3}, value))
{0}		{{
{0}			var e = new System.ComponentModel.CancelEventArgs();
{0}			OnBefore{2}ValueChanging(value, e);
{0}			if (!e.Cancel)
{0}			{{
{0}				{3} = value;
{0}				OnBefore{2}RaisePropertyChanged();
{0}				RaisePropertyChanged({4}Properties.{2});
{0}				OnAfter{2}RaisePropertyChanged();
{0}			}}
{0}		}}
{0}	}}
{0}}}";
		private string PropertyFormat(string indent, string returnType, string propertyName, string variableName, string className)
		{
			return string.Format(_propertyFormat, indent, returnType, propertyName, variableName, className);
		}
	}
}