using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HotelRatesExcelGenerator.Test
{
	using System.IO;
	using Newtonsoft.Json.Linq;

	[TestClass]
	public class HotelRatesExcelGeneratorTest
	{
		[TestMethod]
		public void StringTest()
		{
			var jsonString = File.ReadAllText("hotelrates.json");
			Assert.IsTrue(HotelRatesExcelGenerator.CreateExcelFile(jsonString));
		}

		[TestMethod]
		public void FileInfoTest()
		{
			FileInfo fileInfo = new FileInfo("hotelrates.json");
			Assert.IsTrue(HotelRatesExcelGenerator.CreateExcelFile(fileInfo));
		}

		[TestMethod]
		public void JObjectTest()
		{
			JObject jObject = JObject.Parse(File.ReadAllText("hotelrates.json"));
			Assert.IsTrue(HotelRatesExcelGenerator.CreateExcelFile(jObject));
		}

		[TestMethod]
		public void StreamTest()
		{
			Stream stream = File.OpenRead("hotelrates.json");
			Assert.IsTrue(HotelRatesExcelGenerator.CreateExcelFile(stream));
		}
	}
}
