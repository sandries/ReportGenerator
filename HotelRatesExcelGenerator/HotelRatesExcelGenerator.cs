namespace HotelRatesExcelGenerator
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using Newtonsoft.Json;
	using Newtonsoft.Json.Linq;

	public class HotelRatesExcelGenerator
	{
		public static bool CreateExcelFile(string json)
		{
			dynamic doc = JsonConvert.DeserializeObject<dynamic>(json);
			return InternalCreateExcelFile(doc);
		}

		public static bool CreateExcelFile(Stream stream)
		{
			using (var streamReader = new StreamReader(stream))
			{
				using (var reader = new JsonTextReader(streamReader))
				{
					var jObject = (JObject)JToken.ReadFrom(reader);
					return InternalCreateExcelFile((dynamic)jObject);
				}
			}
		}

		public static bool CreateExcelFile(JObject json)
		{
			return InternalCreateExcelFile((dynamic)json);
		}

		public static bool CreateExcelFile(FileInfo fileinfo)
		{
			using (StreamReader file = File.OpenText(fileinfo.FullName))
			{
				using (var reader = new JsonTextReader(file))
				{
					var jObject = (JObject)JToken.ReadFrom(reader);
					return InternalCreateExcelFile((dynamic)jObject);
				}
			}
		}

		private static bool InternalCreateExcelFile(dynamic json)
		{
			var rates = new List<Rate>();
			foreach (var hotelRate in json.hotelRates)
			{
				var rate = new Rate
				{
					ARRIVAL_DATE = hotelRate.targetDay,
					DEPARTURE_DATE = Convert.ToDateTime(hotelRate.targetDay).AddDays((double)hotelRate.los),
					PRICE = Convert.ToDecimal(hotelRate.price.numericFloat),
					CURRENCY = hotelRate.price.currency,
					RARENAME = hotelRate.rateName,
					ADULTS = hotelRate.adults,
					BREAKFAST_INCLUDED = CheckIfBreakfastIsIncluded(hotelRate)
				};
				rates.Add(rate);
			}

			string fileName = CreateFileName(json);
			rates.ToExcelDocument(fileName);

			return true;
		}

		private static int CheckIfBreakfastIsIncluded(dynamic hotelRate)
		{
			int breakfstIncluded = 0;
			foreach (var rateTag in hotelRate.rateTags)
			{
				if (Convert.ToBoolean(rateTag.shape) && rateTag.name.ToString().ToLower().Equals("breakfast"))
				{
					breakfstIncluded = 1;
					break;
				}
			}

			return breakfstIncluded;
		}

		private static string CreateFileName(dynamic json)
		{
			return @"D:\test" + @"\" + json.hotel.hotelID + "_" + json.hotel.name + ".xlsx";
		}
	}
}
