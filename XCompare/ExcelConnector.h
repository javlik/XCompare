#pragma once
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "Cnterior.h"
#include "Msg.h"

/**
 * @brief Encapsulates OLE Automation access to one Excel workbook (one compared table).
 *
 * CChildView creates two instances — @c m_excel1 and @c m_excel2 — one per table.
 * The shared @c CApplication (Excel process singleton) is kept in CChildView and
 * passed by reference to methods that need it.
 */
class ExcelConnector
{
public:
	ExcelConnector()
		: m_covOptional(COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR))
	{}

	/**
	 * @brief Opens an Excel file and makes it visible.
	 *
	 * Creates or reuses the shared @c Excel.Application singleton passed in @p app.
	 * @param path Full path of the .xlsx/.xls file to open.
	 * @param app  Shared application object; created automatically if not yet connected.
	 * @return @c true on success, @c false if Excel could not be started.
	 */
	bool openFile(const CString& path, CApplication& app)
	{
		if (!app)
		{
			if (!app.CreateDispatch(TEXT("Excel.Application")))
			{
				AfxMessageBox(CMsg(IDS_EXCEL_CANNOT_RUN));
				return false;
			}
		}
		m_Books = app.get_Workbooks();
		m_Book  = m_Books.Open(path,
			m_covOptional, m_covOptional, m_covOptional, m_covOptional,
			m_covOptional, m_covOptional, m_covOptional, m_covOptional,
			m_covOptional, m_covOptional, m_covOptional, m_covOptional,
			m_covOptional, m_covOptional);
		app.put_Visible(TRUE);
		app.put_UserControl(TRUE);
		m_Sheets   = m_Book.get_Worksheets();
		m_filename = path;
		return isOpen();
	}

	/**
	 * @brief Closes the current workbook. Safe to call when no workbook is open.
	 */
	void closeBook()
	{
		if (isOpen())
		{
			try
			{
				m_Book.Close(m_covOptional, m_covOptional, m_covOptional);
			}
			catch (COleException*) {}
		}
	}

	/**
	 * @brief Selects a sheet by name and loads its used range into internal safe arrays.
	 * @param sheetName Name of the worksheet tab to select.
	 * @param outRows   Receives the 1-based number of rows in the used range.
	 * @param outCols   Receives the 1-based number of columns in the used range.
	 */
	void selectSheet(const CString& sheetName, long& outRows, long& outCols)
	{
		m_saRet.Destroy();
		m_Sheet = m_Sheets.get_Item(COleVariant(sheetName));
		m_Range = m_Sheet.get_UsedRange();
		m_saRet = m_Range.get_Value(m_covOptional);
		m_saRet.GetUBound(1, &outRows);
		m_saRet.GetUBound(2, &outCols);
		m_saTmpRet.Destroy();
		m_saTmpRet = m_Range.get_Value(m_covOptional);
	}

	/**
	 * @brief Returns the cell value at (@p column, @p row) from the in-memory safe array.
	 *
	 * Both @p column and @p row are 1-based. Returns an empty string on any OLE error.
	 */
	CString getCellValue(int column, int row)
	{
		long index[2];
		COleVariant vData;
		CString szdata;
		index[0] = row;
		index[1] = column;
		try
		{
			m_saRet.GetElement(index, vData);
			vData = (CString)vData;
		}
		catch (COleException*)
		{
			vData = L"";
		}
		szdata = vData;
		return szdata;
	}

	/**
	 * @brief Returns the cell value from the secondary (temporary) safe array.
	 *
	 * Used by key-suggestion routines that run concurrently with the main comparison pass.
	 * Both @p column and @p row are 1-based.
	 */
	CString getTmpCellValue(int column, int row)
	{
		long index[2];
		COleVariant vData;
		CString szdata;
		index[0] = row;
		index[1] = column;
		try
		{
			m_saTmpRet.GetElement(index, vData);
			vData = (CString)vData;
		}
		catch (COleException*)
		{
			vData = L"";
		}
		szdata = vData;
		return szdata;
	}

	/**
	 * @brief Fills a cell range with a solid background colour via OLE Automation.
	 * @param startCell Top-left cell in Excel A1 notation (e.g. @c "B3").
	 * @param endCell   Bottom-right cell in Excel A1 notation.
	 * @param color     RGB colour value (Windows @c COLORREF).
	 */
	void markCellRange(const CString& startCell, const CString& endCell, COLORREF color)
	{
		CRange range = m_Sheet.get_Range(COleVariant(startCell), COleVariant(endCell));
		m_Interior   = range.get_Interior();
		m_Interior.put_Color(COleVariant(long(color)));
	}

	/**
	 * @brief Selects and activates a cell so Excel scrolls to it.
	 * @param cellRef Target cell in Excel A1 notation.
	 */
	void selectAndActivateCell(const CString& cellRef)
	{
		CRange range = m_Sheet.get_Range(COleVariant(cellRef), COleVariant(cellRef));
		m_Sheet.Activate();
		range.Select();
	}

	/** @brief Returns @c true if a workbook is currently open. */
	bool isOpen() const { return m_Book.m_lpDispatch != nullptr; }

	/** @brief Returns the full path of the currently open file, or empty string if none. */
	CString getFilename() const { return m_filename; }

	/**
	 * @brief Returns the worksheets collection for populating the sheet combo box.
	 *
	 * Valid only after a successful @c openFile() call.
	 */
	CWorksheets& getSheets() { return m_Sheets; }

private:
	CWorkbooks    m_Books;
	CWorkbook     m_Book;
	CWorksheets   m_Sheets;
	CWorksheet    m_Sheet;
	CRange        m_Range;
	COleSafeArray m_saRet;     ///< Primary safe array (main comparison pass).
	COleSafeArray m_saTmpRet;  ///< Secondary safe array (key-suggestion threads).
	Cnterior      m_Interior;
	COleVariant   m_covOptional;
	CString       m_filename;
};
