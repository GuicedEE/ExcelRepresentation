package com.guicedee.services.excelrepresentation;


import com.fasterxml.jackson.databind.ObjectMapper;
import com.guicedee.services.jsonrepresentation.IJsonRepresentation;
import lombok.extern.java.Log;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;
import java.util.logging.Level;
import static java.math.BigDecimal.*;

/**
 * Reader and simple writer for Excel workbooks backed by Apache POI.
 * Supports legacy {@code .xls} and modern {@code .xlsx} formats via the
 * appropriate workbook implementation and exposes convenience methods for
 * row/column access and object mapping.
 *
 * @author Ernst
 * 		Created:23 Oct 2013
 */
@SuppressWarnings({"WeakerAccess", "unused"})
@Log
public class ExcelReader
		implements AutoCloseable
{
	private InputStream inputStream;

	private HSSFWorkbook oldStyle;
	private XSSFWorkbook xwb;
	private boolean isH;
	private Sheet currentSheet;

	/**
	 * Creates a reader for the first sheet in the provided input stream.
	 *
	 * @param inputStream
	 * 		the Excel file input stream
	 * @param extension
	 * 		the file extension, {@code xls} or {@code xlsx}
	 *
	 * @throws ExcelRenderingException
	 * 		if the stream is null or the workbook cannot be opened
	 */
	public ExcelReader(InputStream inputStream, String extension) throws ExcelRenderingException
	{
		this(inputStream, extension, 0);
	}

	/**
	 * Creates a reader for the given sheet index in the provided input stream.
	 *
	 * @param inputStream
	 * 		the Excel file input stream
	 * @param extension
	 * 		the file extension, {@code xls} or {@code xlsx}
	 * @param sheet
	 * 		the zero-based sheet index to use as the current sheet
	 *
	 * @throws ExcelRenderingException
	 * 		if the stream is null or the workbook cannot be opened
	 */
	public ExcelReader(InputStream inputStream, String extension, int sheet) throws ExcelRenderingException
	{
		if (inputStream == null)
		{
			throw new ExcelRenderingException("Inputstream for document is null");
		}
		this.inputStream = inputStream;
		if (extension.equalsIgnoreCase("xls"))
		{
			try
			{
				oldStyle = new HSSFWorkbook(inputStream);
			}
			catch (Throwable e)
			{
				log.log(Level.SEVERE,"Unable to excel ",e);
				throw new ExcelRenderingException("Cannot open xls workbook",e);
			}
			this.currentSheet = oldStyle.getSheetAt(sheet);
			isH = true;
		}
		else
		{
			try
			{
				xwb = new XSSFWorkbook(inputStream);
			}
			catch (Throwable e)
			{
				log.log(Level.SEVERE,"Unable to excel ",e);
				throw new ExcelRenderingException("Cannot open xlsx workbook", e);
			}
			this.currentSheet = xwb.getSheetAt(sheet);
			isH = false;
		}
	}

	/**
	 * Returns the underlying workbook instance.
	 *
	 * @return the Apache POI workbook
	 */
	public Workbook getWorkbook()
	{
		if (isH)
		{
			return oldStyle;
		}
		else
		{
			return xwb;
		}
	}

	/**
	 * Writes a header row at index {@code 0} in the current sheet.
	 *
	 * @param headers
	 * 		the header values to write
	 */
	public void writeHeader(List<String> headers)
	{
		Row row = currentSheet.createRow(0);
		int counter = 0;
		for (String item : headers)
		{
			Cell cell = row.createCell(counter);
			cell.setCellValue(item);
			counter++;
		}
	}

	/**
	 * Fetches a sheet's data into a rectangular table.
	 *
	 * @param sheetNumber
	 * 		the sheet number, starting at 0
	 * @param start
	 * 		how many rows to skip before collecting
	 * @param records
	 * 		the number of rows to return
	 *
	 * @return a 2D array of values, with row and column counts derived from the sheet
	 */
	public Object[][] fetchRows(int sheetNumber, int start, int records)
	{
		int totalSheetRows = getRowCount(sheetNumber) + 1;
		int totalRowColumns = getColCount(sheetNumber);
		if (records > totalSheetRows)
		{
			records = totalSheetRows;
		}
		int arraySize = records - start;
		if (arraySize == 0)
		{
			arraySize = 1;
		}
		else if (arraySize < 0)
		{
			arraySize = arraySize * -1;
		}
		Object[][] tableOut = new Object[arraySize][getColCount(sheetNumber)];
		Sheet sheet;
		if (isH)
		{
			sheet = oldStyle.getSheetAt(sheetNumber);
		}
		else
		{
			sheet = xwb.getSheetAt(sheetNumber);
		}
		int rowN = 0;
		int cellN = 0;
		int skip = 0;
		try
		{
			for (Row row : sheet)
			{
				if (skip < start)
				{
					skip++;
					continue;
				}
				cellN = 0;
				int tCs = getColCount(sheetNumber);
				for (int cn = 0; cn < tCs; cn++)
				{
					Cell cell = row.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					CellType cellType = cell.getCellType();
					switch (cellType)
					{
						case BLANK:
						{
							tableOut[rowN][cellN] = "";
							break;
						}
						case NUMERIC:
						{
							tableOut[rowN][cellN] = cell.getNumericCellValue();
							if (tableOut[rowN][cellN] instanceof Double)
							{
								Double d = (Double) tableOut[rowN][cellN];
								if (new BigDecimal(d).equals(ZERO))
								{
									tableOut[rowN][cellN] = 0;
								}
							}
							break;
						}
						case STRING:
						{
							tableOut[rowN][cellN] = cell.getStringCellValue();
							break;
						}
						case FORMULA:
						{
							FormulaEvaluator evaluator;
							if (isH)
							{
								evaluator = oldStyle.getCreationHelper()
								               .createFormulaEvaluator();
							}
							else
							{
								evaluator = xwb.getCreationHelper()
								               .createFormulaEvaluator();
							}
							CellValue cellValue = evaluator.evaluate(cell);
							Double valueD = cellValue.getNumberValue();
							tableOut[rowN][cellN] = valueD;
							break;
						}
						case BOOLEAN:
						{
							tableOut[rowN][cellN] = cell.getBooleanCellValue();
							break;
						}
						default:
						{
							break;
						}
					}
					cellN++;
					if (totalRowColumns == cellN)
					{
						break;
					}
				}
				rowN++;
				if (records == rowN)
				{
					break;
				}
			}
		}
		catch (ArrayIndexOutOfBoundsException e)
		{
			log.log(Level.WARNING, "Reached the end of the file before hitting the max results. logic error.", e);
		}
		catch (Exception e)
		{
			log.log(Level.WARNING, "Couldn't go through the whole excel file - ", e);
		}
		return tableOut;
	}

	/**
	 * Returns the number of columns in a sheet based on the row at the same index.
	 *
	 * @param sheetNo
	 * 		the sheet index
	 *
	 * @return the number of cells in the row at {@code sheetNo}
	 */
	public int getColCount(int sheetNo)
	{
		Sheet sheet;
		if (isH)
		{
			sheet = oldStyle.getSheetAt(sheetNo);
		}
		else
		{
			sheet = xwb.getSheetAt(sheetNo);
		}

		Row row = sheet.getRow(sheetNo);
		return row.getLastCellNum();
	}

	/**
	 * Returns the number of rows in the given sheet.
	 *
	 * @param sheetNo
	 * 		the sheet index
	 *
	 * @return the number of rows in the sheet
	 */
	public int getRowCount(int sheetNo)
	{
		if (isH)
		{
			return oldStyle.getSheetAt(sheetNo)
			          .getLastRowNum() + 1;
		}
		else
		{
			return xwb.getSheetAt(sheetNo)
			          .getLastRowNum() + 1;
		}
	}

	/**
	 * Writes a row of strings to the given row index in the current sheet.
	 *
	 * @param rowNumber
	 * 		the row index to create
	 * @param headers
	 * 		the values to write
	 */
	public void writeRow(int rowNumber, List<String> headers)
	{
		Row row = currentSheet.createRow(rowNumber);
		int counter = 0;
		for (String item : headers)
		{
			Cell cell = row.createCell(counter);
			cell.setCellValue(item);
			counter++;
		}
	}

	/**
	 * Writes a row of values to the given row index in the current sheet.
	 * Values are coerced based on their runtime type.
	 *
	 * @param rowNumber
	 * 		the row index to create
	 * @param rowData
	 * 		the values to write
	 */
	@SuppressWarnings("ConstantConditions")
	public void writeRow(int rowNumber, Object[] rowData)
	{
		Row row = currentSheet.createRow(rowNumber);
		int counter = 0;
		for (Object item : rowData)
		{
			Cell cell = row.createCell(counter);
			if (item == null)
			{
				item = "";
			}
			String ftype = item.getClass()
			                   .getName();
			if (ftype.equals("java.lang.String"))
			{
				cell.setCellValue((String) item);
			}
			else if (ftype.equals("java.lang.Boolean") || ftype.equals("boolean"))
			{
				cell.setCellValue((Boolean) item);
			}
			else if (ftype.equals("java.util.Date"))
			{
				cell.setCellValue((Date) item);
			}
			else if (ftype.equals("int") || ftype.equals("java.lang.Integer"))
			{
				cell.setCellValue((Integer) item);
			}
			else if (ftype.equals("long") || ftype.equals("java.lang.Long") || ftype.equals("java.math.BigInteger"))
			{
				cell.setCellValue((Long) item);
			}
			else if (ftype.equals("java.math.BigDecimal"))
			{
				cell.setCellValue(((BigDecimal) item).doubleValue());
			}
			else
			{
				cell.setCellValue((Double) item);
			}
			counter++;
		}
	}

	/**
	 * Returns a row from the current sheet.
	 *
	 * @param rowNumber
	 * 		the row index
	 *
	 * @return the row, or {@code null} if not present
	 */
	public Row getRow(int rowNumber)
	{
		return currentSheet.getRow(rowNumber);
	}

	/**
	 * Returns a cell from the current sheet.
	 *
	 * @param rowNumber
	 * 		the row index
	 * @param cellNumber
	 * 		the cell index within the row
	 *
	 * @return the cell, or {@code null} if not present
	 */
	public Cell getCell(int rowNumber, int cellNumber)
	{
		return currentSheet.getRow(rowNumber)
		                   .getCell(cellNumber);
	}

	/**
	 * Serializes the workbook into a byte array.
	 *
	 * @return the workbook bytes, or {@code null} if an error occurs
	 */
	public byte[] get()
	{
		byte[] output = null;
		try (ByteArrayOutputStream baos = new ByteArrayOutputStream())
		{
			if (isH)
			{
				oldStyle.write(baos);
			}
			else
			{
				xwb.write(baos);
			}
			output = baos.toByteArray();
		}
		catch (Exception e)
		{
			log.log(Level.SEVERE, "Unable to get the byte array for the excel file", e);
		}
		return output;
	}

	/**
	 * Reads rows from the named sheet and maps them into objects using Jackson.
	 * The first row is treated as a header row and used as JSON field names.
	 *
	 * @param sheetName
	 * 		the sheet name to read
	 * @param type
	 * 		the target type to deserialize into
	 * @param <T>
	 * 		the target type
	 *
	 * @return a list of mapped objects
	 */
	@SuppressWarnings("unchecked")
	public <T> List<T> getRecords(String sheetName, Class<T> type)
	{
		int sheetLocation = getWorkbook().getSheetIndex(sheetName);
		Object[][] rows = this.fetchRows(sheetLocation, 0, getRowCount(sheetLocation));
		Object[] headerRow = rows[0];
		List<T> output = new ArrayList<>();
		Map<Integer, Map<String, String>> cells = new TreeMap<>();
		for (int i = 1; i < rows.length; i++)
		{
			//Cell <Header,Value>
			cells.put(i, new LinkedHashMap<>());
			JSONObject rowData = new JSONObject();
			for (int j = 0; j < getColCount(sheetLocation); j++)
			{
				if (rows[i][j] instanceof BigDecimal)
				{
					rowData.put(headerRow[j].toString(), ((BigDecimal) rows[i][j]).toPlainString());
				}
				else
				{
					rowData.put(headerRow[j].toString(), rows[i][j])
					       .toString();
				}
				if(rows[i][j] != null)
				{
					cells.get(i)
					     .put(headerRow[j].toString(), rows[i][j].toString());
				}
				else
				{
					//end of row?
				}
				//cells.put(i, rows[i][j].toString().trim());
			}
			String outcome = rowData.toString();
			try
			{
				ObjectMapper om = IJsonRepresentation.getObjectMapper();
				T typed = om.readValue(outcome, type);
				output.add(typed);
			}
			catch (Exception e)
			{
				log.log(Level.SEVERE, "Unable to build an object from the references - " + outcome, e);
			}
		}
		return output;
	}

	/**
	 * Closes the underlying workbook and input stream.
	 *
	 * @throws Exception
	 * 		if closing fails
	 */
	@Override
	public void close() throws Exception
	{
		try
		{
			if (isH)
			{
				oldStyle.close();
			}
			else
			{
				xwb.close();
			}

		}
		catch (Exception e)
		{
			log.log(Level.SEVERE, "Unable to write the excel file out", e);
		}
		inputStream.close();
	}


}
