using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace YearbookVerify {
	public partial class MainWindow : Window {
		//members
		private bool input;
		private Spelling spellCheck;
		private List<Name> names;
		//constructors
		public MainWindow() {
			InitializeComponent();
			input = true;
			spellCheck = new Spelling();
			names = new List<Name>();
			WordDictionary wd = new WordDictionary();
			wd.DictionaryFile = "registeredNames.dic";
			wd.Add("cat");
			wd.Add("dog");
			wd.Add("cut");
			spellCheck.Dictionary = wd;
			spellCheck.SuggestionMode = Spelling.SuggestionEnum.PhoneticNearMiss;
			spellCheck.Suggest("dig");
			foreach (string str in spellCheck.Suggestions)
				Console.WriteLine(str); 
		}
		//methods
		private string FindSomeSuggestions(string word) {
			spellCheck.Suggest(word);
			List<string> suggs = new List<string>();
			for (int i = 0; i < spellCheck.Suggestions.Count && i < 3; i++) {
				suggs.Add(spellCheck.Suggestions[i] as string);
			}
			string sugg = string.Join(" or ", suggs);
			return sugg.Length > 0 ? " Did you mean: " + sugg + "?" : "";
		}
		private string MakeError(string s) {
			return "***" + s + "***";
		}
		private SpellingResult CheckSpelling(string text, char delim) {
			string[] lines = text.Split(delim);
			int errors = 0;
			int noExist = 0;
			string markedLines = "";
			int lineNum = 1;
			foreach(string s in lines) {
				string editLine = s.Trim();
				if (editLine.Length > 0) {
					string[] words = editLine.Split(' ');
					if(words.Length == 2) {
						bool sFirst = spellCheck.TestWord(words[0]);
						bool sLast = spellCheck.TestWord(words[1]);
						if(sFirst && sLast) {
							Name n = names.Find(name => name.First.Equals(words[0]) && name.Last.Equals(words[1]));
							if(n == null) {
								markedLines += MakeError(editLine) + " This name combination does not exist." + delim;
								noExist++;
							}
							else {
								markedLines += editLine + delim;
                            }
						}
						else {
							if (sFirst) {
								markedLines += words[0] + " ";
							}
							else {
								markedLines += MakeError(words[0]) + " " + FindSomeSuggestions(words[0]) + " ";
								errors++;
							}
							if (sLast) {
								markedLines += words[1] + delim;
                            }
							else {
								markedLines += MakeError(words[1]) + " " + FindSomeSuggestions(words[1]) + delim;
								errors++;
							}
						}
					}
					else {
						return new SpellingResult("Error on line " + lineNum + ". Name must be two words.", -1);
					}
				}
				lineNum++;
			}
			markedLines = markedLines.Substring(0, markedLines.Length - 1);
			return new SpellingResult(markedLines, errors);
		}
		//wpf
		private void actionButton_Click(object sender, RoutedEventArgs e) {
			if (input) {
				if (inputBox.Text.Length > 0) {
					//verify
					SpellingResult res = CheckSpelling(inputBox.Text, '\n');
					if (res.UserError) {
						MessageBox.Show(res.UserMessage);
						return;
					}
					else {
						inputBox.Text = res.MarkedLines;
					}
				}
			}
			else {
				//return
			}
			actionButton.Content = input ? "Return" : "Verify";
			input = !input;
			inputBox.IsEnabled = !inputBox.IsEnabled;
		}
	}
	class SpellingResult {
		//members
		private int errors;
		private string markedLines;
		//constructors
		internal SpellingResult(string markedLines, int errors) {
			this.errors = errors;
			this.markedLines = markedLines;
		}
		//properties
		internal string MarkedLines { get { return markedLines; } }
		internal bool NoErrors { get { return errors == 0; } }
		internal int Errors { get { return errors; } }
		internal bool UserError { get { return errors == -1; } }
		internal string UserMessage { get { return markedLines; } }
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
