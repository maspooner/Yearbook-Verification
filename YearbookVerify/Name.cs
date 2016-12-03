using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YearbookVerify {
	/// <summary>
	/// Models a first and last name pairing.
	/// </summary>
	public class Name {
		//constants
		private const char BASE64_SEPERATOR = '%';
		//properties
		public string First { get; private set; } //given name
		public string Last  { get; private set; } //family name
		//constructors
		public Name(string first, string last) {
			First = first;
			Last = last;
		}
		//methods
		/// <summary>
		/// Encodes this data structure as a Base64 string. To Decode, use <see cref="Decode"/>
		/// </summary>
		public string ToEncoded() {
			return Convert.ToBase64String(Encoding.UTF8.GetBytes(First)) 
					+ BASE64_SEPERATOR
					+ Convert.ToBase64String(Encoding.UTF8.GetBytes(Last));
		}
		public override string ToString() {
			return First + " " + Last;
		}
		public bool IsSame(Name n) {
			return n.First.Equals(First) && n.Last.Equals(Last);
		}
		//statics
		/// <summary>
		/// Parses a line of words into a <seealso cref="Name"/> structure.
		/// Supports 2 to 4 words
		/// </summary>
		/// <param name="s">the line to parse</param>
		public static Name Parse(string s) {
			s = s.Trim();
			//remove superfluous whitespace
			while(s.Contains("  ")) {
				s = s.Replace("  ", " ");
            }
			//split by ' '
			string[] parts = s.Split(' ');
			string fir = "", las = "";
			//ex: Ava Maria St. Pierre
			if(parts.Length == 4) {
				fir = parts[0] + " " + parts[1];
                las = parts[2] + " " + parts[3];
			}
			//ex: Ava St. Pierre
			else if(parts.Length == 3) {
				//ambiguous first/last name pairing for second part
				//Contains a .? Probably a last name matching
				if(parts[1].IndexOf('.') != -1) {
					fir = parts[0];
					las = parts[1] + " " + parts[2];
				}
				//Probably a first name matching
				else {
					fir = parts[0] + " " + parts[1];
					las = parts[2];
				}
			}
			//ex: Matthew Spooner
			else if (parts.Length == 2) {
				fir = parts[0];
				las = parts[1];
			}
			else {
				//invalid name
				return null;
			}
			//make captials as they should be
			return new Name(FixupCapitals(fir), FixupCapitals(las));
		}
		/// <summary>
		/// Captializes the right letters of a name string.
		/// ex: ana maria -> Ana Maria
		/// ex: st. Pierre-marx -> St. Pierre-Marx
		/// ex: matt -> Matt
		/// </summary>
		public static string FixupCapitals(string s) {
			char nameDelim = ' ';
			if (s.IndexOf(' ') == -1) {
				if(s.IndexOf('-') == -1) {
					//no special capitalization, just cap first letter
					return char.ToUpperInvariant(s[0]) + s.Substring(1).ToLowerInvariant();
				}
				else {
					nameDelim = '-';
				}
			}
			//capitalize both portions of the name
			return FixupCapitals(s.Substring(0, s.IndexOf(nameDelim))) + nameDelim
				 + FixupCapitals(s.Substring(s.IndexOf(nameDelim) + 1));
		}
		/// <summary>
		/// Decodes a previously encoded <seealso cref="Name"/> structure from Base64.
		/// </summary>
		/// <param name="encoded">the Base64 encoded string</param>
		public static Name Decode(string encoded) {
			string[] names = encoded.Split(BASE64_SEPERATOR);
			return new Name(Encoding.UTF8.GetString(Convert.FromBase64String(names[0])),
							Encoding.UTF8.GetString(Convert.FromBase64String(names[1])));
		}
	}
}
