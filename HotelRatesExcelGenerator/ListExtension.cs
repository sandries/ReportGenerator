namespace HotelRatesExcelGenerator
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.Reflection;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Spreadsheet;

	public static class ListExtension
	{
		public static DataTable ToDataTable<T>(this List<T> list)
		{
			var dt = new DataTable();

			foreach (PropertyInfo info in typeof(T).GetProperties())
			{
				dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
			}

			foreach (T t in list)
			{
				DataRow row = dt.NewRow();
				foreach (PropertyInfo info in typeof(T).GetProperties())
				{
					if (!IsNullableType(info.PropertyType))
					{
						row[info.Name] = info.GetValue(t, null);
					}
					else
					{
						row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
					}
				}
				dt.Rows.Add(row);
			}
			return dt;
		}

		private static Type GetNullableType(Type t)
		{
			Type returnType = t;
			if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
			{
				returnType = Nullable.GetUnderlyingType(t);
			}
			return returnType;
		}

		private static bool IsNullableType(Type type)
		{
			return (type == typeof(string) || type.IsArray || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>)));
		}

		public static void ToExcelDocument<T>(this List<T> list, string fileFullName)
		{
			var ds = new DataSet();
			ds.Tables.Add(list.ToDataTable());

			using (var document = SpreadsheetDocument.Create(fileFullName, SpreadsheetDocumentType.Workbook))
			{
				WriteExcelFile(ds, document);
			}
		}

		private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
		{
			spreadsheet.AddWorkbookPart();
			spreadsheet.WorkbookPart.Workbook = new Workbook();

			spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

			uint worksheetNumber = 1;
			foreach (DataTable dt in ds.Tables)
			{
				//  For each worksheet you want to create
				var newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
				newWorksheetPart.Worksheet = new Worksheet();

				// create sheet data
				newWorksheetPart.Worksheet.AppendChild(new SheetData());

				// save worksheet
				WriteDataTableToExcelWorksheet(dt, newWorksheetPart);
				newWorksheetPart.Worksheet.Save();

				// create the worksheet to workbook relation
				if (worksheetNumber == 1)
				{
					spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
				}

				spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet
				{
					Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
					SheetId = worksheetNumber,
					Name = dt.TableName
				});

				worksheetNumber++;
			}

			spreadsheet.WorkbookPart.Workbook.Save();
		}

		private static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart)
		{
			var worksheet = worksheetPart.Worksheet;
			var sheetData = worksheet.GetFirstChild<SheetData>();

			int numberOfColumns = dt.Columns.Count;
			bool[] isNumericColumn = new bool[numberOfColumns];

			string[] excelColumnNames = new string[numberOfColumns];
			for (int n = 0; n < numberOfColumns; n++)
			{
				excelColumnNames[n] = GetExcelColumnName(n);
			}

			uint rowIndex = 1;

			var headerRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
			sheetData.Append(headerRow);

			for (int colInx = 0; colInx < numberOfColumns; colInx++)
			{
				DataColumn col = dt.Columns[colInx];
				AppendTextCell(excelColumnNames[colInx] + "1", col.ColumnName, headerRow);
				isNumericColumn[colInx] = (col.DataType.FullName == "System.Decimal") || (col.DataType.FullName == "System.Int32");
			}

			foreach (DataRow dr in dt.Rows)
			{
				++rowIndex;
				var newExcelRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
				sheetData.Append(newExcelRow);

				for (int colInx = 0; colInx < numberOfColumns; colInx++)
				{
					var cellValue = dr.ItemArray[colInx].ToString();

					if (isNumericColumn[colInx])
					{
						double cellNumericValue = 0;
						if (double.TryParse(cellValue, out cellNumericValue))
						{
							cellValue = cellNumericValue.ToString();
							AppendNumericCell(excelColumnNames[colInx] + rowIndex, cellValue, newExcelRow);
						}
					}
					else
					{
						AppendTextCell(excelColumnNames[colInx] + rowIndex, cellValue, newExcelRow);
					}
				}
			}
		}

		private static string GetExcelColumnName(int columnIndex)
		{
			if (columnIndex < 26)
			{
				return ((char)('A' + columnIndex)).ToString();
			}

			char firstChar = (char)('A' + (columnIndex / 26) - 1);
			char secondChar = (char)('A' + (columnIndex % 26));

			return $"{firstChar}{secondChar}";
		}

		private static void AppendTextCell(string cellReference, string cellStringValue, Row excelRow)
		{
			var cell = new Cell { CellReference = cellReference, DataType = CellValues.String };
			cell.Append(new CellValue { Text = cellStringValue });
			excelRow.Append(cell);
		}

		private static void AppendNumericCell(string cellReference, string cellStringValue, Row excelRow)
		{
			var cell = new Cell { CellReference = cellReference };
			cell.Append(new CellValue { Text = cellStringValue });
			excelRow.Append(cell);
		}
	}
}