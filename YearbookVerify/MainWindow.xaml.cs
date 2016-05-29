using Microsoft.Office.Interop.Excel;
using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;

namespace YearbookVerify {
	public partial class MainWindow : System.Windows.Window {
		//members
		private Spelling spellCheck;
		private List<Name> names;
		//constructors
		public MainWindow() {
			//TODO handle exceptions
			InitializeComponent();
			spellCheck = new Spelling();
			names = CompileMatchedList(); //TODO load from file
			WordDictionary wd = new WordDictionary();
			wd.DictionaryFile = "registeredNames.dic";
			spellCheck.Dictionary = wd;
			spellCheck.SuggestionMode = Spelling.SuggestionEnum.PhoneticNearMiss;
			foreach (string str in spellCheck.Suggestions)
				Console.WriteLine(str);
			File.WriteAllText("registeredNames.dic", Properties.Resources.dicBase + string.Join(Environment.NewLine, FindUnique()));
		}
		//methods
		private List<Name> CompileMatchedList() {
			List<Name> compNames = new List<Name>();
			Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
			app.Visible = false;
			Workbook wb = app.Workbooks.Open(@"F:\Documents\Visual Studio 2015\Projects\YearbookVerify\YearbookVerify\bin\Debug\input.xlsx");
			Worksheet ws = wb.Worksheets[1];
			string first, last;
			int i = 3;
			bool done;
			do {
				first = (string)ws.Cells[i, 2].Value;
				last = (string)ws.Cells[i, 3].Value;
				done = first == null || last == null || first.Length == 0 || last.Length == 0;
				if (!done) {
					compNames.Add(new Name(first, last));
					Console.WriteLine(first + " " + last);
					i++;
				}
            }
			while (!done);
			wb.Close(false);
			app.Quit();
			ReleaseInterop(ws);
			ReleaseInterop(wb);
			ReleaseInterop(app);
			return compNames;
		}
		private List<string> FindUnique() {
			List<string> unique = new List<string>();
			foreach(Name n in names) {
				if (!unique.Contains<string>(n.First))
					unique.Add(n.First);
				if (!unique.Contains<string>(n.Last))
					unique.Add(n.Last);
			}
			foreach (string s in unique)
				Console.WriteLine(s);
			return unique;
		}
		private void ReleaseInterop(object obj) {
			try {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex) {
				obj = null;
				MessageBox.Show("Unable to release the Object " + ex.ToString());
			}
			finally {
				GC.Collect();
			}
		}
		private string FindSomeSuggestions(string word) {
			spellCheck.Suggest(word);
			List<string> suggs = new List<string>();
			for (int i = 0; i < spellCheck.Suggestions.Count && i < 3; i++) {
				suggs.Add(CapFirst(spellCheck.Suggestions[i] as string));
			}
			string sugg = string.Join(" or ", suggs);
			return sugg.Length > 0 ? " Did you mean: " + sugg + "?" : "";
		}
		private string MakeError(string s) {
			return "***" + s + "***";
		}
		private string CapFirst(string s) {
			return char.ToUpperInvariant(s[0]) + s.Substring(1).ToLowerInvariant();
		}
		private SpellingResult CheckSpelling(string text, char delim) {
			string[] lines = text.Split(delim);
			int errors = 0;
			int noExist = 0;
			string[] regLines = new string[lines.Length];
			string[] markedLines = new string[lines.Length];
			for(int i = 0; i < lines.Length; i++) {
				string editLine = lines[i].Trim();
				if (editLine.Length > 0) {
					string[] words = editLine.Split(' ');
					if(words.Length == 2) {
						words[0] = CapFirst(words[0]);
						words[1] = CapFirst(words[1]);
						bool sFirst = spellCheck.TestWord(words[0]);
						bool sLast = spellCheck.TestWord(words[1]);
						regLines[i] = words[0] + " " + words[1];
                        if (sFirst && sLast) {
							Name n = names.Find(name => name.First.Equals(words[0]) && name.Last.Equals(words[1]));
							if(n == null) {
								markedLines[i] = "***This name combination not registered.***";
								noExist++;
							}
							else {
								markedLines[i] = "All good.";
                            }
						}
						else {
							if(!sFirst) {
								markedLines[i] = MakeError(words[0]) + FindSomeSuggestions(words[0]);
								errors++;
							}
							if(!sLast) {
								if (markedLines[i].Length > 0)
									markedLines[i] += "\t";
								markedLines[i] += MakeError(words[1]) + FindSomeSuggestions(words[1]);
								errors++;
							}
						}
					}
					else {
						return new SpellingResult("Error on line " + (i+1) + ". Name must be two words.");
					}
				}
				else {
					regLines[i] = "";
					markedLines[i] = "All good.";
				}
			}
			return new SpellingResult(regLines, markedLines, errors);
		}
		//wpf
		private void actionButton_Click(object sender, RoutedEventArgs e) {
			if (inputBox.Text.Length > 0) {
				const char delim = '\n';
				//verify
				SpellingResult res = CheckSpelling(inputBox.Text, delim);
				if (res.UserError) {
					MessageBox.Show(res.UserMessage);
					return;
				}
				else {
					inputBox.Text = string.Join(delim.ToString(), res.RegLines);
					outputBox.Text = string.Join(delim.ToString(), res.MarkedLines);
				}
			}
		}
		private void reloadButton_Click(object sender, RoutedEventArgs e) {
			names = CompileMatchedList();
			File.WriteAllText("registeredNames.dic", Properties.Resources.dicBase + string.Join(Environment.NewLine, FindUnique()));
		}
		private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e) {
			ScrollViewer v = sender as ScrollViewer;
			if(sender == view1) {
				view2.ScrollToVerticalOffset(view1.VerticalOffset);
			}
			else {
				view1.ScrollToVerticalOffset(view2.VerticalOffset);
			}
		}
	}
	class SpellingResult {
		//members
		private int errors;
		private string[] regLines;
		private string[] markedLines;
		private string userMessage;
		//constructors
		internal SpellingResult(string[] regLines, string[] markedLines, int errors) {
			this.errors = errors;
			this.regLines = regLines;
			this.markedLines = markedLines;
			userMessage = null;
		}
		internal SpellingResult(string userMessage) {
			this.userMessage = userMessage;
			errors = -1;
			markedLines = null;
		}
		//properties
		internal string[] RegLines { get { return regLines; } }
		internal string[] MarkedLines { get { return markedLines; } }
		internal bool NoErrors { get { return errors == 0; } }
		internal int Errors { get { return errors; } }
		internal bool UserError { get { return errors == -1; } }
		internal string UserMessage { get { return userMessage; } }
	}
	class Name {
		//members
		private string first;
		private string last;
		//constructors
		internal Name(string first, string last) {
			this.first = first;
			this.last = last;
		}
		//properties
		internal string First { get { return first; } }
		internal string Last { get { return last; } }
	}
}
