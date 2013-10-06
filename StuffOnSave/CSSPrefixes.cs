using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using EnvDTE;
using EnvDTE80;

namespace PrefixCSS
{
	internal class CSSPrefixes
	{
		private const string CleanCssSuffix = ".clean.css";
		private const string PrefixedCssSuffix = ".prefixed.css";
		private static readonly Regex rxPrefixedLine = new Regex(@"^\s*-(?:ms|moz|webkit|o)-|-(?:ms|moz|webkit|o)-calc|/\*\s*calc\s+fallback\s*\*/", RegexOptions.Compiled);
		private static readonly Regex rxPrefixedAtKeyframes = new Regex(@"^\s*@-(?:ms|moz|webkit|o)-keyframes\b", RegexOptions.Compiled);
		private static readonly Regex rxUnprefixedAtKeyframes = new Regex(@"^\s*@keyframes\b", RegexOptions.Compiled);
		private static readonly Regex rxCalc = new Regex(@"\b(?<!\-)calc\((?<inner>.+?)\)\s*;", RegexOptions.Compiled);
		private static readonly Regex rxPercentUnits = new Regex(@"[\-\d\.]+%", RegexOptions.Compiled);

		private static readonly Regex rxPropertiesToWebkitPrefix =
			new Regex(@"\b(?<!\-)(?<keyword>keyframes|transform|transition|animation|user-select|font-feature-settings|box-sizing)\b", RegexOptions.Compiled);
		private static readonly Regex rxPropertiesToMozPrefix = new Regex(
			@"\b(?<!\-)(?<keyword>keyframes|transform|transition|animation|user-select|font-feature-settings|box-sizing)\b", RegexOptions.Compiled);
		private static readonly Regex rxPropertiesToMsPrefix =
			new Regex(@"\b(?<!\-)(?<keyword>transform|user-select|font-feature-settings)\b", RegexOptions.Compiled);
		
		private readonly OutputWindowPanes _outputWindowPanes;
		private OutputWindowPane _ourOutputWindowPane;
		private string _currentFilePath;

		public CSSPrefixes(DTE2 dte)
		{
			_outputWindowPanes = dte.ToolWindows.OutputWindow.OutputWindowPanes;
		}

		private void OutputMessage(string fmt, params object[] args)
		{
			if (_ourOutputWindowPane == null)
			{
				_ourOutputWindowPane = _outputWindowPanes.Add("CSS Vendor Prefixer");
			}

			_ourOutputWindowPane.OutputString(string.Format(fmt, args) + "\r\n");
		}

		private void OutputWarning(int line, string fmt, params object[] args)
		{
			OutputMessage("Line {0} in {1}: {2}", line, Path.GetFileName(_currentFilePath), string.Format(fmt, args));
		}

		public void Add(string filepath)
		{
			if (!File.Exists(filepath))
			{
				OutputMessage("{0} does not exist!", filepath);
				return;
			}

			_currentFilePath = filepath;

			var originalLastWriteTime = File.GetLastWriteTime(filepath);
			var prefixedFilePath = filepath + PrefixedCssSuffix;

			if (File.Exists(prefixedFilePath))
			{
				var prefixedLastWriteTime = File.GetLastWriteTime(prefixedFilePath);
				if (prefixedLastWriteTime == originalLastWriteTime)
				{
					OutputMessage("{0} is up-to-date", Path.GetFileName(prefixedFilePath));
					return;
				}

				if (prefixedLastWriteTime > originalLastWriteTime)
				{
					OutputMessage("{1} is newer than {0}! Creating cleaned version of prefixed file instead.",
								  Path.GetFileName(filepath), Path.GetFileName(prefixedFilePath));
					ReadLinesAndClean(prefixedFilePath);
					return;
				}
			}

			List<string> lines;
			if (ReadLinesAndClean(filepath, out lines))
				return; // returns true on error

			var preAddCount = lines.Count;

			AddVendorPrefix("-ms-", lines, rxPropertiesToMsPrefix);
			AddVendorPrefix("-moz-", lines, rxPropertiesToMozPrefix);
			AddVendorPrefix("-webkit-", lines, rxPropertiesToWebkitPrefix);

			if (lines.Count != preAddCount)
			{
				File.WriteAllLines(prefixedFilePath, lines, Encoding.UTF8);
				File.SetLastWriteTime(prefixedFilePath, originalLastWriteTime);
				OutputMessage("{2} vendor-prefixed lines added to {0} to create {1}", Path.GetFileName(filepath), Path.GetFileName(prefixedFilePath), lines.Count - preAddCount);
			}
			else
			{
				OutputMessage("No need for vendor-prefixing in {0}", Path.GetFileName(filepath));
			}
		}

		private void ReadLinesAndClean(string filepath)
		{
			List<string> lines;
			ReadLinesAndClean(filepath, out lines);
		}

		private bool ReadLinesAndClean(string filepath, out List<string> lines)
		{
			lines = new List<string>();

			try
			{
				lines = new List<string>(File.ReadAllLines(filepath, Encoding.UTF8));
			}
			catch (IOException ioex)
			{
				OutputMessage("Error '{0}' reading {1}!", ioex.Message, filepath);
				return true;
			}

			var preCleanCount = lines.Count;

			ClearExistingPrefixes(lines);

			if (lines.Count != preCleanCount)
			{
				OutputMessage("{1} vendor-prefixed lines found in {0}", Path.GetFileName(filepath), preCleanCount - lines.Count);
				var cleanFilePath = (filepath + CleanCssSuffix).Replace(PrefixedCssSuffix + CleanCssSuffix, CleanCssSuffix);
				File.WriteAllLines(cleanFilePath, lines, Encoding.UTF8);
				OutputMessage("Created {0}", Path.GetFileName(cleanFilePath));
			}

			return false;
		}

		private void AddVendorPrefix(string prefix, IList<string> lines, Regex rxPropertyToPrefix)
		{
			var insertedLines = -1;	// this results in adding a line number to make zero become one

			for (var i = 0; i < lines.Count; ++i)
			{
				var lineNumber = i - insertedLines;

				if (rxPrefixedAtKeyframes.IsMatch(lines[i]))
				{
					//  skip the already-prefixed keyframes block
					i += CollectBlock(lines, i).Count - 1; // -1 because of the for-loop increment
					continue;
				}

				if (rxUnprefixedAtKeyframes.IsMatch(lines[i]))
				{
					var blockLines = CollectBlock(lines, i);

					if (prefix != "-ms-")
					{
						//  insert each line of the collected block prefixing the relevant keywords
						foreach (var line in blockLines)
						{
							lines.Insert(i, rxPropertyToPrefix.Replace(line, match => prefix + match.Value));
							++i;
						}
						insertedLines += blockLines.Count;
					}

					//  skip over the unprefixed keyframes block
					i += blockLines.Count - 1; // -1 because of the for-loop increment
					continue;
				}

				if (rxCalc.IsMatch(lines[i]))
				{
					if (prefix != "-ms-")
					{
						//  add a prefixed calc() for all but -ms-
						lines.Insert(i, rxCalc.Replace(lines[i], match => prefix + match.Value));
						++i; // adjust for inserted line
						insertedLines += 1;
					}
					else
					{
						var noCalc = rxCalc.Replace(lines[i], delegate(Match match)
																	 {
																		 var exp = match.Groups["inner"].Value;

																		 var percentUnitsMatch = rxPercentUnits.Match(exp);
																		 if (percentUnitsMatch.Success)
																		 {
																			 OutputWarning(lineNumber, "Questionable fallback of '{1};' added for '{0}'", match.Value, percentUnitsMatch.Value);
																			 return percentUnitsMatch.Value + ";";
																		 }

																		 exp = exp.Replace("px", "");
																		 try
																		 {
																			 var result = Convert.ToDouble(new DataTable().Compute(exp, null));
																			 var roundedResult = Math.Round(result, 0);

																			 if (result < roundedResult || result > roundedResult)
																				 OutputWarning(lineNumber, "'{0}' resulted in fractional value {1}; fallback rounded to {2}", match.Value, result, roundedResult);

																			 return string.Format("{0}px;", roundedResult);
																		 }
																		 catch (SyntaxErrorException)
																		 {
																			 OutputWarning(lineNumber, "Could not compute '{0}'; no fallback added", match.Value);
																			 return match.Value;
																		 }
																	 }
							);

						if (noCalc != lines[i])
						{
							//  add a fallback value when we're doing -ms-
							lines.Insert(i, noCalc + " /* calc fallback */");
							++i; // adjust for inserted line
							insertedLines += 1;
						}
					}
				}

				if (rxPropertyToPrefix.IsMatch(lines[i]))
				{
					lines.Insert(i, rxPropertyToPrefix.Replace(lines[i], match => prefix + match.Value));
					++i; // adjust for inserted line
					insertedLines += 1;
				}
			}
		}

		private void ClearExistingPrefixes(IList<string> lines)
		{
			//	remove the single-line rules
			for (var i = 0; i < lines.Count; ++i)
			{
				if (!rxPrefixedLine.IsMatch(lines[i]))
					continue;

				OutputMessage("Removed '{0}'", lines[i].Trim());

				lines.RemoveAt(i);
				--i; // adjust for removed line
			}

			//  remove the prefixed @keyframes blocks
			for (var i = 0; i < lines.Count; ++i)
			{
				if (!rxPrefixedAtKeyframes.IsMatch(lines[i]))
					continue;

				OutputMessage("Removed '{0}'", lines[i].Trim());

				var linesToDelete = CollectBlock(lines, i).Count;

				while (linesToDelete > 0)
				{
					lines.RemoveAt(i);
					--linesToDelete;
				}

				--i; // adjust for removed line
			}
		}

		private static List<string> CollectBlock(IList<string> lines, int i)
		{
			var blockLines = new List<string>();

			var braces = 0;

			//	find the opening brace
			while (braces == 0)
			{
				blockLines.Add(lines[i]);

				if (lines[i].Contains("{"))
					braces = 1;

				++i;
			}

			while (braces > 0)
			{
				blockLines.Add(lines[i]);

				if (lines[i].Contains("{"))
					++braces;

				if (lines[i].Contains("}"))
					--braces;

				++i;
			}

			return blockLines;
		}
	}
}