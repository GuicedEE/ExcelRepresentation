package com.guicedee.services.excelrepresentation;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Provides convenience methods for reading Excel content into objects and
 * exporting object data to Excel formats.
 */
public interface IExcelRepresentation
{

	/**
	 * Reads an Excel file into its object representation using the first row
	 * as headers for field mapping.
	 *
	 * @param stream
	 * 		the Excel input stream
	 * @param objectType
	 * 		the type to map each row into
	 * @param sheetName
	 * 		the sheet name to read from
	 * @param <T>
	 * 		the target type
	 *
	 * @return a list of mapped objects
	 *
	 * @throws ExcelRenderingException
	 * 		if the workbook cannot be read or parsed
	 */
	default <T> List<T> fromExcel(InputStream stream, Class<T> objectType, String sheetName) throws ExcelRenderingException
	{
		try (ExcelReader excelReader = new ExcelReader(stream, "xlsx"))
		{
			return excelReader.getRecords(sheetName, objectType);
		}
		catch (IOException e)
		{
			throw new ExcelRenderingException("Cannot read the excel file");
		}
		catch (Exception e)
		{
			throw new ExcelRenderingException("General error with the excel file", e);
		}
	}

	/**
	 * Converts the implementing type into an Excel representation.
	 *
	 * @return a string representation of the Excel content, or {@code null} if not implemented
	 */
	default String toExcel()
	{
		return null;
	}


}
