#pragma once
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "Cnterior.h"
#include "Msg.h"

/// <summary>
/// Encapsulates OLE Automation access to one Excel workbook (one compared table).
/// CChildView creates two instances - m_excel1 and m_excel2 - one per table.
/// The shared CApplication (Excel process singleton) is kept in CChildView and
/// passed by reference to methods that need it.
/// </summary>
class ExcelConnector
{
public:
	ExcelConnector()
		: m_covOptional(COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR))
	{}

	/// <summary>
	/// Opens an Excel file. Returns true on success.
	/// Creates or reuses the shared Excel.Application singleton passed in.
	/// </summary>
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

	/// <summary>
	/// Closes the current workbook. Safe to call when no workbook is open.
	/// </summary>
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

	/// <summary>
	/// Selects a sheet by name and loads its used range into the internal safe arrays.
	/// outRows and outCols receive the 1-based dimensions of the used range.
	/// </summary>
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

	/// <summary>
	/// Returns the cell value at (column, row) from the in-memory safe array.
	/// column and row are 1-based. Returns an empty string on error.
	/// </summary>
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

	/// <summary>
	/// Returns the cell value from the secondary (temporary) safe array.
	/// Used by key-suggestion routines that run concurrently with the main pass.
	/// </summary>
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

	/// <summary>
	/// Marks a contiguous range of cells with the given RGB colour.
	/// startCell and endCell use Excel A1 notation (e.g. "B3").
	/// </summary>
	void markCellRange(const CString& startCell, const CString& endCell, COLORREF color)
	{
		CRange range = m_Sheet.get_Range(COleVariant(startCell), COleVariant(endCell));
		m_Interior   = range.get_Interior();
		m_Interior.put_Color(COleVariant(long(color)));
	}

	/// <summary>
	/// Selects and activates a cell in the sheet so Excel scrolls to it.
	/// cellRef uses Excel A1 notation. Call CApplication methods separately
	/// to bring Excel to the foreground if needed.
	/// </summary>
	void selectAndActivateCell(const CString& cellRef)
	{
		CRange range = m_Sheet.get_Range(COleVariant(cellRef), COleVariant(cellRef));
		m_Sheet.Activate();
		range.Select();
	}

	/// <summary>Returns true if a workbook is currently open.</summary>
	bool isOpen() const { return m_Book.m_lpDispatch != nullptr; }

	/// <summary>Returns the path of the currently open file (empty if none).</summary>
	CString getFilename() const { return m_filename; }

	/// <summary>
	/// Provides access to the worksheets collection for sheet-combo population.
	/// Valid after a successful openFile() call.
	/// </summary>
	CWorksheets& getSheets() { return m_Sheets; }

private:
	CWorkbooks    m_Books;
	CWorkbook     m_Book;
	CWorksheets   m_Sheets;
	CWorksheet    m_Sheet;
	CRange        m_Range;
	COleSafeArray m_saRet;
	COleSafeArray m_saTmpRet;
	Cnterior      m_Interior;
	COleVariant   m_covOptional;
	CString       m_filename;
};
