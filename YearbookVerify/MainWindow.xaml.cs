using Microsoft.Office.Interop.Excel;
using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace YearbookVerify {
	public partial class MainWindow : System.Windows.Window, IDisposable {
		//members
		private Spelling spellCheck;
		private List<Name> names;
		//constructors
		public MainWindow() {
			InitializeComponent();
			Icon = GetIcon(Properties.Resources.icon);
			spellCheck = new Spelling();
			names = LoadNames();
			WordDictionary wd = new WordDictionary();
			wd.DictionaryFile = "registeredNames.dic";
			spellCheck.Dictionary = wd;
			spellCheck.SuggestionMode = Spelling.SuggestionEnum.PhoneticNearMiss;
		}
		//methods
		private BitmapSource GetIcon(Bitmap b) {
			BitmapImage bitmapImage;
            using (MemoryStream memory = new MemoryStream()) {
				b.Save(memory, ImageFormat.Bmp);
				memory.Position = 0;
				bitmapImage = new BitmapImage();
				bitmapImage.BeginInit();
				bitmapImage.StreamSource = memory;
				bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
				bitmapImage.EndInit();
			}
			return bitmapImage;
		}
		private void ReloadDataStructures() {
			names = CompileNames();
			SaveNames();
			WriteDictionary();
			spellCheck.Dictionary.Dispose();
			WordDictionary wd = new WordDictionary();
			wd.DictionaryFile = "registeredNames.dic";
			spellCheck.Dictionary = wd;
		}
		private void SaveNames() {
			try {
				File.WriteAllText(Environment.CurrentDirectory + "/matchedNames.nms", string.Join("\n", names));
			}
			catch (Exception e) {
				MessageBox.Show("An error occured attempting to save valid names to file.\n\nError:\t" + e.Message, "Error Writing Names File", MessageBoxButton.OK);
			}
		}
		private List<Name> LoadNames() {
			try {
				List<Name> names = new List<Name>();
				string[] ls = File.ReadAllLines(Environment.CurrentDirectory + "/matchedNames.nms");
				foreach(string l in ls) {
					Console.WriteLine(l);
					string[] words = l.Split(' ');
					names.Add(new Name(words[0], words[1]));
				}
				return names;
			}
			catch (Exception e) {
				MessageBox.Show("An error occured attempting to load valid names from file.\n\nError:\t" + e.Message, "Error Reading Names File", MessageBoxButton.OK);
			}
			return new List<Name>();
		}
		private void WriteDictionary() {
			try {
				File.WriteAllText(Environment.CurrentDirectory + "/registeredNames.dic", Properties.Resources.dicBase + string.Join(Environment.NewLine, FindUnique()));
			}
			catch (Exception e) {
				MessageBox.Show("An error occured trying to convert the input spreadsheet into a dictionary file.\n\nError:\t" + e.Message, "Error Writing Dictionary File", MessageBoxButton.OK);
			}
		}
		private List<Name> CompileNames() {
			try {
				List<Name> compNames = new List<Name>();
				Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
				app.Visible = false;
				Workbook wb = app.Workbooks.Open(Environment.CurrentDirectory + "/input.xlsx");
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
			catch(Exception e) {
				MessageBox.Show("An error occured when trying to read the input spreadsheet. Make sure the file \'input.xlsx\' is located in the directory" +
					"with this program and that it is formatted properly. The first first name should be in cell B3 and the first last name should be in cell"+
					" C3.\n\nError:\t" + e.Message, "Error reading spreadsheet", MessageBoxButton.OK);
			}
			return new List<Name>();
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
				MessageBox.Show("An error has occured: " + ex.Message);
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
		private string CapFirst(string s) {
			return char.ToUpperInvariant(s[0]) + s.Substring(1).ToLowerInvariant();
		}
		private string Substring(string s, int i, int j) {
			return s.Substring(i, j - i);
		}
		private SpellingResult CheckSpelling(string text, char delim) {
			string[] lines = text.Split(delim);
			List<string> regLines = new List<string>();
			string markedLines = "";
			int lineNum = 0;
			int i = 1;
			foreach(string l in lines) {
				string editLine = l.Trim();
				while (editLine.Contains("  ")) {
					int iSpace = editLine.IndexOf("  ");
					editLine = Substring(editLine, 0, iSpace) + Substring(editLine, iSpace + 1, editLine.Length);
				}
				if (editLine.Length > 0) {
					string[] words = editLine.Split(' ');
					if(words.Length == 2) {
						words[0] = CapFirst(words[0]);
						words[1] = CapFirst(words[1]);
						bool sFirst = spellCheck.TestWord(words[0]);
						bool sLast = spellCheck.TestWord(words[1]);
						regLines.Add(words[0] + " " + words[1]);
						lineNum++;
						if (sFirst && sLast) {
							Name n = names.Find(name => name.First.Equals(words[0]) && name.Last.Equals(words[1]));
							if(n == null) {
								markedLines += "Entry " + lineNum + ": \"" + regLines[regLines.Count - 1] + "\" name is not registered." + "\n";
							}
						}
						else {
							if(!sFirst) {
								markedLines += "Entry " + lineNum +": \"" + words[0] + "\" is incorrectly spelled. " + FindSomeSuggestions(words[0]) + "\n";
							}
							if(!sLast) {
								markedLines += "Entry " + lineNum + ": \"" + words[1] + "\" is incorrectly spelled. " + FindSomeSuggestions(words[1]) + "\n";
							}
						}
					}
					else {
						return new SpellingResult("Error on entry " + i + ". Name must be two words.");
					}
				}
				i++;
			}
			return new SpellingResult(regLines.ToArray(), markedLines);
		}
		public void Dispose() {
			spellCheck.Dispose();
		}
		//wpf
		private void actionButton_Click(object sender, RoutedEventArgs e) {
			if (inputBox.Text.Length > 0) {
				//verify
				SpellingResult res = CheckSpelling(inputBox.Text, ',');
				if (res.UserError) {
					MessageBox.Show(this, res.UserMessage, "User Error");
					return;
				}
				else {
					inputBox.Text = string.Join(", ", res.RegLines);
					
					outputBox.Text = res.MarkedLines.Length == 0 ? "All good." : string.Join("\n", res.MarkedLines);
				}
			}
		}
		private void reloadButton_Click(object sender, RoutedEventArgs e) {
			ReloadDataStructures();
		}
		private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e) {
			ScrollViewer v = sender as ScrollViewer;
			if (sender == view1) {
				view2.ScrollToVerticalOffset(view1.VerticalOffset);
			}
			else {
				view1.ScrollToVerticalOffset(view2.VerticalOffset);
			}
		}
	}
	class SpellingResult {
		//members
		private string[] regLines;
		private string markedLines;
		private string userMessage;
		//constructors
		internal SpellingResult(string[] regLines, string markedLines) {
			this.regLines = regLines;
			this.markedLines = markedLines;
			userMessage = null;
		}
		internal SpellingResult(string userMessage) {
			this.userMessage = userMessage;
			regLines = null;
			markedLines = null;
		}
		//properties
		internal string[] RegLines { get { return regLines; } }
		internal string MarkedLines { get { return markedLines; } }
		internal bool UserError { get { return markedLines == null; } }
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
		public override string ToString() {
			return first + " " + last;
		}
	}
}
