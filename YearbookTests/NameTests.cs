using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using YearbookVerify;

namespace YearbookTests {
	[TestClass]
	public class NameTests {
		//test data
		private const string RAW_NAME_1 = "  matthew   Spooner  ";
		private static Name NAME_1 = new Name("Matthew", "Spooner");
		//test methods
		private string VerboseName(Name n) {
			return "F:(" + n.First + ") L:(" + n.Last + ")";
		}
		private void AssertSameName(Name expected, Name actual) {
			Console.WriteLine("Expected: " + VerboseName(expected));
			Console.WriteLine("Actual: "   + VerboseName(actual));
			Assert.IsTrue(expected.IsSame(actual));
		}
		[TestMethod]
		public void EncodeDecode_EqualToOriginal() {
			AssertSameName(NAME_1, Name.Decode(NAME_1.ToEncoded()));
		}
		[TestMethod]
		public void Parse_TwoWordNameNoSpecial() {
			AssertSameName(NAME_1, Name.Parse(RAW_NAME_1));
		}
		[TestMethod]
		public void Parse_TwoWordNameHasSpecial() {
			AssertSameName(new Name("Ana", "Pierre-Quanto"),
				Name.Parse("  ana    pierre-Quanto    "));
		}
		[TestMethod]
		public void Parse_ThreeWordNameTwoFirstNamesNoSpecial() {
			AssertSameName(new Name("Ana Maria", "Pierre"),
				Name.Parse("  ana   Maria  pierre    "));
		}
		[TestMethod]
		public void Parse_ThreeWordNameTwoLastNamesNoSpecial() {
			AssertSameName(new Name("Ana", "St. Pierre"),
				Name.Parse("  ana   st.  pierre    "));
		}
		[TestMethod]
		public void Parse_ThreeWordNameTwoFirstNamesHasSpecial() {
			AssertSameName(new Name("Ana Maria", "Pierre-Quanto"),
				Name.Parse("  ana   Maria  pierre-quanto    "));
		}
		[TestMethod]
		public void Parse_ThreeWordNameTwoLastNamesHasSpecial() {
			AssertSameName(new Name("Ana", "St. Pierre-Quanto"), 
				Name.Parse("  ana   st.  pierre-quanto    "));
		}
		[TestMethod]
		public void Parse_FourWordNameNoSpecial() {
			AssertSameName(new Name("Ana Maria", "St. Pierre"), 
				Name.Parse("  ana  maria  st.  pierre    "));
		}
		[TestMethod]
		public void Parse_FourWordNameHasSpecial() {
			AssertSameName(new Name("Ana Maria", "St. Pierre-Maxwell"), 
				Name.Parse("  ana  maria  st.  pierre-maxwell    "));
		}
		[TestMethod]
		public void Parse_TooShortIsNull() {
			Assert.IsNull(Name.Parse("  ana  "));
		}
		[TestMethod]
		public void Parse_TooLongIsNull() {
			Assert.IsNull(Name.Parse("  ana maria st. martino vanchez "));
		}
	}
}
