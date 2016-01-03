namespace HotelRatesExcelGenerator
{
	using System;

	class Rate
	{
		public DateTime ARRIVAL_DATE { get; set; }

		public DateTime DEPARTURE_DATE { get; set; }

		public decimal PRICE { get; set; }

		public string CURRENCY { get; set; }

		public string RARENAME { get; set; }

		public int ADULTS { get; set; }

		public int BREAKFAST_INCLUDED { get; set; }
	}
}