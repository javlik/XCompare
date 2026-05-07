// This file declares the CChildView class, which is responsible for the main application window's client area.
// ChildView.h : interface of the CChildView class
//


#pragma once
#include "Constants.h"
#include "TableData.h"
#include "ComparisonMatrix.h"
#include "ExcelConnector.h"
#include "ComparisonEngine.h"
#include "KeyFinder.h"
#include <vector>
#include <map>
#include <atomic>

/**
 * @brief Main client-area window of the XCompare application.
 *
 * Owns the two @c ExcelConnector instances, the @c ComparisonEngine, the @c KeyFinder,
 * all ribbon UI element pointers, the comparison matrix and the scroll/paint state.
 * Nearly all user interactions (ribbon commands, mouse events, worker-thread notifications)
 * are handled here.
 */
class CChildView : public CWnd
{
// Construction
public:
	CChildView();

// Attributes
public:

// Operations
public:

// Overrides
	protected:
	/** @brief Adjusts the window class before creation (registers custom window class). */
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);

// Implementation
public:
	virtual ~CChildView();

	// Generated message map functions
protected:
	/** @brief Handles WM_PAINT: allocates GDI resources then delegates to the paint helpers. */
	afx_msg void OnPaint();
	DECLARE_MESSAGE_MAP()

	/**
	 * @brief Groups all temporary GDI objects and visible-area bounds for a single paint pass.
	 *
	 * Created on the stack inside @c OnPaint() and passed to each paint helper by reference.
	 */
	struct PaintCtx
	{
		CPen   *pen1, *pen2, *pen3, *pen4, *pen5, *pen6,
		       *pen7, *pen8, *pen9, *pen10, *pen11, *pen12;
		CBrush *brush0, *brush1, *brush2, *brush3,
		       *brush4, *brush5, *brush6, *brush7;
		CFont  *font1, *font2, *font3, *font4,
		       *font1B, *font2B, *font1C, *font2C;
		int    bnd_X_min, bnd_X_max, bnd_Y_min, bnd_Y_max; ///< Visible cell-unit bounds.
	};

	/** @brief Draws the top-left info area (file names, status text). */
	void paintInfoArea      (CDC& dc, PaintCtx& ctx);
	/** @brief Draws the grid lines separating matrix cells. */
	void paintGridLines     (CDC& dc, PaintCtx& ctx);
	/** @brief Draws the left row-header column (row numbers / key-missing markers). */
	void paintRowHeaders    (CDC& dc, PaintCtx& ctx);
	/** @brief Draws the top column-header row (column names / colour-coded match state). */
	void paintColumnHeaders (CDC& dc, PaintCtx& ctx);
	/** @brief Draws the interior cells of the comparison matrix (match percentage / colour fill). */
	void paintMatrixCells   (CDC& dc, PaintCtx& ctx);
	/** @brief Draws the curved similarity-line overlay connecting matched columns. */
	void paintSimilarityLines(CDC& dc, PaintCtx& ctx);

public:
	/** @brief Ribbon command: opens a file picker for the first Excel file. */
	afx_msg void OnPickFirstFile();
	/** @brief Ribbon command: opens a file picker for the second Excel file. */
	afx_msg void OnPickSecondFile();
	/**
	 * @brief Shared implementation for picking an Excel file.
	 * @param pNewFile     Set to true if a new file was successfully opened.
	 * @param pExcel       ExcelConnector to open the file into.
	 * @param pTable       Table descriptor to populate after sheet selection.
	 * @param pSheetCombo  Ribbon combo box to fill with sheet names.
	 * @param pFilename    Receives the chosen file path.
	 */
	void pickFile(bool* pNewFile, ExcelConnector* pExcel, Table* pTable, CMFCRibbonComboBox* pSheetCombo, CString* pFilename);
	/** @brief Ribbon command: runs the comparison (builds key arrays and calls firstPass). */
	afx_msg void OnCreateMatrix();
	/** @brief Enables the "Pick first sheet" combo only when a first file is open. */
	afx_msg void OnUpdatePickFirstSheet(CCmdUI *pCmdUI);
	/** @brief Enables the "Compare" button only when both sheets have been selected. */
	afx_msg void OnUpdateCreateMatrix(CCmdUI *pCmdUI);
	/** @brief Updates the file-1 name label in the ribbon. */
	afx_msg void OnUpdateFilename1(CCmdUI *pCmdUI);
	/** @brief Updates the file-2 name label in the ribbon. */
	afx_msg void OnUpdateFilename2(CCmdUI *pCmdUI);
	/**
	 * @brief Shared helper: sets a ribbon label text to the given filename (shortened if needed).
	 * @param pCmdUI     The ribbon label to update.
	 * @param pszFilename The full path to display.
	 * @param idString   String-table ID of a fallback "(no file)" label.
	 */
	void updateFileName(CCmdUI* pCmdUI, CString* pszFilename, int idString);
	/** @brief Translates vertical mouse-wheel rotation into a vertical scroll. */
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	/** @brief Enables the "Pick second sheet" combo only when a second file is open. */
	afx_msg void OnUpdatePickSecondSheet(CCmdUI *pCmdUI);
	/** @brief Updates progress bar 1 and 2 (comparison pass progress). */
	afx_msg void OnUpdateProgress1(CCmdUI *pCmdUI);
	/** @brief Handles vertical scroll-bar events (thumb drag, arrow, page). */
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	/** @brief Handles horizontal scroll-bar events (thumb drag, arrow, page). */
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);

	/**
	 * @brief Shared implementation for selecting a worksheet from a combo box.
	 *
	 * Reads the currently selected sheet name, calls @c ExcelConnector::selectSheet(),
	 * and syncs the spinner and dimension ribbon controls.
	 */
	void pickSheet(ExcelConnector* pExcel, Table* pTable, CMFCRibbonComboBox* pSheetCombo, CMFCRibbonEdit* pSpinner_Names, CMFCRibbonEdit* pSpinner_Fdata, CMFCRibbonEdit* pRows, CMFCRibbonEdit* pCols);
	/** @brief Ribbon command: selects a sheet from the first file. */
	void OnPickFirstSheet();
	/** @brief Ribbon command: increments/decrements the header-row index for table 1. */
	afx_msg void OnSpin1Names();
	/** @brief Enables the header-row spinner for table 1 only when a sheet is loaded. */
	afx_msg void OnUpdateSpin1Names(CCmdUI *pCmdUI);
	/** @brief Enables the first-data-row spinner for table 1 only when a sheet is loaded. */
	afx_msg void OnUpdateSpin1Fdata(CCmdUI *pCmdUI);
	/** @brief Ribbon command: increments/decrements the first-data-row index for table 1. */
	afx_msg void OnSpin1Fdata();
	/** @brief Repopulates the sheet combo box and resets spinners for table 1 after a new file is opened. */
	void updateCombos1();
	/** @brief Ribbon command: selects a sheet from the second file. */
	afx_msg void OnPickSecondSheet();
	/** @brief Repopulates the sheet combo box and resets spinners for table 2 after a new file is opened. */
	void updateCombos2();
	/** @brief Enables the first-data-row spinner for table 2 only when a sheet is loaded. */
	afx_msg void OnUpdateSpin2Fdata(CCmdUI *pCmdUI);
	/** @brief Ribbon command: increments/decrements the first-data-row index for table 2. */
	afx_msg void OnSpin2Fdata();
	/** @brief Enables the header-row spinner for table 2 only when a sheet is loaded. */
	afx_msg void OnUpdateSpin2Names(CCmdUI *pCmdUI);
	/** @brief Ribbon command: increments/decrements the header-row index for table 2. */
	afx_msg void OnSpin2Names();
	/** @brief Handles a left double-click: selects the clicked matrix cell as a key column pair. */
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	/** @brief Runs the main comparison pass on the UI thread (used when both key arrays are ready). */
	void firstPass();
	/** @brief Builds the concatenated key string array and lookup map for table 1. @return 0 = success, 1 = duplicate key. */
	int createKeyArrays1();
	/** @brief Builds the concatenated key string array and lookup map for table 2. @return 0 = success, 2 = duplicate key. */
	int createKeyArrays2();
	/** @brief Verifies key uniqueness for table 1 (O(n²) pass). @return @c true if all keys are unique. */
	bool checkKeysUniqueness1();
	/** @brief Verifies key uniqueness for table 2 (O(n²) pass). @return @c true if all keys are unique. */
	bool checkKeysUniqueness2();
	/** @brief Handles mouse movement: updates the hovered cell and refreshes the status bar. */
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	/** @brief Ribbon command: reads the similarity threshold slider value and triggers a matrix refresh. */
	afx_msg void OnSimilarityThreshold();
	/** @brief Enables the similarity-threshold slider when the comparison matrix exists. */
	afx_msg void OnUpdateSimilarityThreshold(CCmdUI *pCmdUI);
	/** @brief Ribbon command: toggles marking of differences in file 1. */
	afx_msg void OnMarkInFile1();
	/** @brief Enables the "Mark in file 1" checkbox when a comparison result is available. */
	afx_msg void OnUpdateMarkInFile1(CCmdUI *pCmdUI);
	/** @brief Ribbon command: toggles marking of differences in file 2. */
	afx_msg void OnMarkInFile2();
	/** @brief Enables the "Mark in file 2" checkbox when a comparison result is available. */
	afx_msg void OnUpdateMarkInFile2(CCmdUI *pCmdUI);
	/** @brief Ribbon command: launches the key-suggestion worker threads for both tables. */
	afx_msg void OnSuggestKeys();
	/**
	 * @brief Converts a (row, column) cell address to Excel A1 notation (e.g. "B3").
	 * @param row 1-based row index.
	 * @param clm 1-based column index.
	 * @return CString in Excel A1 format.
	 */
	CString convertR1C1(int row, int clm);
	/** @brief Marks cell (@p row, @p clm) in table 1 with the selected highlight colour via OLE. */
	void markIn1(int row, int clm);
	/** @brief Marks cell (@p row, @p clm) in table 2 with the selected highlight colour via OLE. */
	void markIn2(int row, int clm);
	/** @brief Sets up horizontal and vertical scroll bar ranges based on current matrix dimensions. */
	void initScrollBars();
	/** @brief Handles WM_SIZE: updates client size, resets scroll bars and invalidates the window. */
	afx_msg void OnSize(UINT nType, int cx, int cy);
	/** @brief Handles WM_CREATE: caches all ribbon element pointers and initialises state. */
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	/** @brief Updates progress bar 3 (key-building pass progress). */
	afx_msg void OnUpdateProgress2(CCmdUI *pCmdUI);

	/** @brief Enables the "Verify keys" checkbox when a comparison result is available. */
	afx_msg void OnUpdateVerifyKeys(CCmdUI *pCmdUI);
	/** @brief Ribbon command: re-runs the comparison with key uniqueness verification enabled. */
	afx_msg void OnVerifyKeys();
	/** @brief Enables the "Suggest keys" button when both sheets are loaded and no search is running. */
	afx_msg void OnUpdateSuggestKeys(CCmdUI *pCmdUI);

	/** @brief Ribbon command: toggles the "same column names only" restriction for key suggestion. */
	afx_msg void OnSameNamesOnly();
	/** @brief Reflects the current state of the "same names only" toggle onto the ribbon button. */
	afx_msg void OnUpdateSameNamesOnly(CCmdUI *pCmdUI);
protected:
	/** @brief Custom message: updates progress bar 1 with the value in lParam (0–100). */
	afx_msg LRESULT OnCmUpdateProgress(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: updates progress bar 2 with the value in lParam (0–100). */
	afx_msg LRESULT OnCmUpdateProgress2(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: updates progress bar 3 with the value in lParam (0–100). */
	afx_msg LRESULT OnCmUpdateProgress3(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: updates the key-search progress bar 1 with the value in lParam. */
	afx_msg LRESULT OnCmUpdateKeyProgress1(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: updates the key-search progress bar 2 with the value in lParam. */
	afx_msg LRESULT OnCmUpdateKeyProgress2(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: table-1 key-array thread has finished; triggers uniqueness check or firstPass. */
	afx_msg LRESULT OnCmKeys1Done(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: table-2 key-array thread has finished; triggers uniqueness check or firstPass. */
	afx_msg LRESULT OnCmKeys2Done(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: table-1 key-suggestion thread has finished gathering candidates. */
	afx_msg LRESULT OnCmGathering1Done(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: table-2 key-suggestion thread has finished gathering candidates. */
	afx_msg LRESULT OnCmGathering2Done(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: cross-check found a valid key combination; applies it and runs firstPass. */
	afx_msg LRESULT OnCmKeysFound(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: cross-check could not find a valid key combination; notifies the user. */
	afx_msg LRESULT OnCmKeysNotFound(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: marking thread has finished; invalidates the window. */
	afx_msg LRESULT OnCmMarkingReady(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: table-1 column-similarity thread has finished. */
	afx_msg LRESULT OnCmSims1Done(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: table-2 column-similarity thread has finished. */
	afx_msg LRESULT OnCmSims2Done(WPARAM wParam, LPARAM lParam);
	/** @brief Custom message: the firstPass worker thread has finished; refreshes the matrix display. */
	afx_msg LRESULT OnCmFirstPassDone(WPARAM wParam, LPARAM lParam);
public:
	/** @brief Applies background-colour marks to all differing cells in both Excel files. */
	void markInFiles();
	/** @brief Ribbon command: opens the colour picker for highlight colour 1 (differences). */
	afx_msg void OnColorPicker1();
	/** @brief Enables the colour-picker-1 button when a comparison result is available. */
	afx_msg void OnUpdateColorPicker1(CCmdUI *pCmdUI);
	/** @brief Ribbon command: opens the colour picker for highlight colour 2 (matches). */
	afx_msg void OnColorPicker2();
	/** @brief Enables the colour-picker-2 button when a comparison result is available. */
	afx_msg void OnUpdateColorPicker2(CCmdUI *pCmdUI);
	/** @brief Ribbon command: toggles the auto-mark mode (automatically marks on compare). */
	afx_msg void OnAutoMark();
	/** @brief Reflects the current auto-mark toggle state onto the ribbon button. */
	afx_msg void OnUpdateAutoMark(CCmdUI *pCmdUI);
	/** @brief Decides whether to call @c markInFiles() based on the current auto-mark state. */
	void resolveAutoMark();
	/** @brief Processes all pending Windows messages without blocking, used between long operations. */
	void DrainMsgQueue(void);
	/** @brief Ribbon command: opens the list of found differences in the combo box. */
	afx_msg void OnDiffslist();
	/** @brief Enables the differences list when there are differences to navigate. */
	afx_msg void OnUpdateDiffslist(CCmdUI *pCmdUI);
	/** @brief Ribbon command: scrolls Excel file 1 to the currently selected difference row. */
	afx_msg void OnGotoDiffInFile1();
	/** @brief Returns the row index currently selected in the differences combo box. */
	int rowFromCombo();
	/** @brief Ribbon command: scrolls Excel file 2 to the currently selected difference row. */
	afx_msg void OnGotoDiffInFile2();
	/** @brief Ribbon command: brings the Excel window to the foreground. */
	afx_msg void OnBringExcelToFront();
	/** @brief Enables the "Bring Excel to front" button when at least one file is open. */
	afx_msg void OnUpdateBringExcelToFront(CCmdUI *pCmdUI);
	/** @brief Launches the key-suggestion worker thread for table 1. */
	void suggestKeys1();
	/** @brief Launches the key-suggestion worker thread for table 2. */
	void suggestKeys2();
private:
	/** @brief Shell delegating to @c m_keyFinder.createTempKeyArrays1() (body moved to KeyFinder). */
	int createTempKeyArrays1();
	/** @brief Shell delegating to @c m_keyFinder.createTempKeyArrays2() (body moved to KeyFinder). */
	int createTempKeyArrays2();
	/** @brief Resets all candidate key arrays in @c m_keyFinder. */
	void clearPossibleKeys();
	/** @brief Sorts three integers @p a, @p b, @p c in ascending order in-place. */
	void sort3(int & a, int & b, int & c);
	/**
	 * @brief Computes column-similarity scores for a range of table-1 columns against all table-2 columns.
	 * @param c_i1_start  First table-1 column index to process (1-based).
	 * @param c_i1_end    Last table-1 column index to process (inclusive).
	 * @param progressMsg Custom window message to post for progress updates.
	 * @param doneMsg     Custom window message to post when the range is complete.
	 * @param useTmp      If true, reads from the temporary safe array instead of the primary one.
	 */
	void findSimsRange(int c_i1_start, int c_i1_end, UINT progressMsg, UINT doneMsg, bool useTmp);
public:
	/** @brief Cross-checks both candidate key sets and posts @c CM_KEYS_FOUND or @c CM_KEYS_NOT_FOUND. @return @c true if a matching pair was found. */
	bool mutualCheck();
	/** @brief Removes all key pairs from the engine and resets the key counter. */
	void deleteAllKeys();
	/** @brief Handles a right-click: shows a context menu for the hovered matrix cell. */
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	/** @brief Applies the best candidate key pair from @c m_keyFinder into @c m_engine. @return @c true if keys were applied. */
	bool usePossibleKeys();
	/** @brief Returns the number of column indices in the best candidate key pair. */
	int getNumberOfPossibleKeys();
	/** @brief Enables the key-search-complexity spin control when no search is running. */
	afx_msg void OnUpdateKeySearchComplexity(CCmdUI *pCmdUI);
	/** @brief Ribbon command: reads the complexity value from the spin control and applies it to @c m_keyFinder. */
	afx_msg void OnKeySearchComplexity();
	/**
	 * @brief Returns the number of non-zero key slots in candidate @p item for @p table up to @p order slots.
	 * @param table 1 or 2.
	 * @param order Maximum number of slots to count.
	 * @param item  Index into the PossibleKeys array.
	 */
	int getNumberOfPossibleKeys(int table, int order, int item);
	/** @brief Launches the similarity-computation threads for both tables (low-RAM fallback path). */
	void findSims();
	/** @brief Launches the similarity-computation worker thread for table 1. */
	void findSims1();
	/** @brief Launches the similarity-computation worker thread for table 2. */
	void findSims2();
	/** @brief Ribbon command: toggles display of the column-similarity overlay lines. */
	afx_msg void OnShowSimilarColumns();
	/** @brief Enables the "Show similar columns" button when similarity data is available. */
	afx_msg void OnUpdateShowSimilarColumns(CCmdUI *pCmdUI);
	/** @brief Ribbon command: starts the column-relation detection (similarity computation). */
	afx_msg void OnFindColumnRelations();
	/** @brief Ribbon command: creates or rebuilds the key-index structures in @c m_engine. */
	afx_msg void OnIdxcrtBtn();
	/** @brief Updates the key-search progress bar 1 in the ribbon. */
	afx_msg void OnUpdateKeyProgress1(CCmdUI *pCmdUI);
	/** @brief Updates the key-search progress bar 2 in the ribbon. */
	afx_msg void OnUpdateKeyProgress2(CCmdUI *pCmdUI);
	/** @brief Called when both similarity threads are done: sorts results and refreshes the display. */
	void finishFindRelations();
	/** @brief Enables the "Use key indexing" checkbox (currently always enabled). */
	afx_msg void OnUpdateIdxCheckbox(CCmdUI *pCmdUI);
	/**
	 * @brief Searches @p lpszData backwards from @p startpos for @p lpszSub.
	 * @return 0-based position of the last occurrence, or -1 if not found.
	 */
	int ReverseFind(LPCTSTR lpszData, LPCTSTR lpszSub, int startpos);
	/** @brief Ribbon command (legacy/unused): toggles the key-index checkbox. */
	afx_msg void OnCheckIdx();
	/** @brief Reflects the key-index checkbox state onto the ribbon (legacy/unused). */
	afx_msg void OnUpdateCheckIdx(CCmdUI *pCmdUI);
	/** @brief Ribbon command: toggles the use of index-based key disambiguation. */
	afx_msg void OnUseKeyIndexing();
	/** @brief Reflects the "Use key indexing" toggle state onto the ribbon button. */
	afx_msg void OnUpdateUseKeyIndexing(CCmdUI *pCmdUI);
	/** @brief Enables the row-count edit for table 1 when a sheet is loaded. */
	afx_msg void OnUpdateRows1(CCmdUI *pCmdUI);
	/** @brief Ribbon command: applies the manually entered row count for table 1. */
	afx_msg void OnRows1();
	/** @brief Enables the column-count edit for table 1 when a sheet is loaded. */
	afx_msg void OnUpdateCols1(CCmdUI *pCmdUI);
	/** @brief Ribbon command: applies the manually entered column count for table 1. */
	afx_msg void OnCols1();
	/** @brief Enables the row-count edit for table 2 when a sheet is loaded. */
	afx_msg void OnUpdateRows2(CCmdUI *pCmdUI);
	/** @brief Ribbon command: applies the manually entered row count for table 2. */
	afx_msg void OnRows2();
	/** @brief Ribbon command: applies the manually entered column count for table 2. */
	afx_msg void OnCols2();
	/** @brief Enables the column-count edit for table 2 when a sheet is loaded. */
	afx_msg void OnUpdateCols2(CCmdUI *pCmdUI);

	// ---------------------------------------------------------------
	// Data members (moved from .cpp file-scope globals)
	// ---------------------------------------------------------------

	// Public: accessed directly by worker thread functions
public:
	std::atomic<bool>         m_bUniqueKeys1{false}; ///< @c true when the key combination is unique in table 1.
	std::atomic<bool>         m_bUniqueKeys2{false}; ///< @c true when the key combination is unique in table 2.
	std::atomic<bool>         m_bLockPrg1{false};    ///< @c true while the table-1 key-build thread is running.
	std::atomic<bool>         m_bLockPrg2{false};    ///< @c true while the table-2 key-build thread is running.
	ComparisonEngine m_engine;    ///< Main comparison engine (key arrays, maps, firstPass result).
	KeyFinder        m_keyFinder; ///< Entropy-based key-suggestion engine.

private:
	// Synchronisation between threads
	std::atomic<bool> m_bWaitingForKeys{false};       ///< @c true when firstPass is blocked waiting for both key arrays.
	std::atomic<bool> m_bKeys1done{false};            ///< @c true when the table-1 key-array thread has finished.
	std::atomic<bool> m_bKeys2done{false};            ///< @c true when the table-2 key-array thread has finished.
	std::atomic<bool> m_bKeysGathering1done{false};   ///< @c true when the table-1 key-suggestion thread has finished.
	std::atomic<bool> m_bKeysGathering2done{false};   ///< @c true when the table-2 key-suggestion thread has finished.

	CString m_szRsltTxt;       ///< Formatted result text shown in the info area after a comparison.

	Palette m_Palette[20];     ///< Array of 20 user-selectable highlight colours.

	int  m_nOldx;              ///< Previous hovered cell column (for dirty-region optimisation).
	int  m_nOldy;              ///< Previous hovered cell row.
	int  m_nChosenColor1;      ///< Index into @c m_Palette for highlight colour 1 (differences).
	int  m_nChosenColor2;      ///< Index into @c m_Palette for highlight colour 2 (matches).
	Clnt m_Clnt;               ///< Cached client area rectangle.
	bool m_bDoAutoMark;        ///< If @c true, cells are automatically marked after each comparison.
	int  m_nMatrixDone;        ///< Non-zero when the comparison matrix has been computed.
	int  m_nPrereqDone;        ///< Non-zero when both prerequisite passes have completed.
	bool m_bMarkIdentCols;     ///< If @c true, fully-identical columns are also colour-marked.
	bool m_bSameNames;         ///< If @c true, only columns with the same name are compared.
	int  m_nEffMax;            ///< Effective maximum scroll position (columns × cell width).

	// Dynamic data arrays
	std::vector<bool>    m_pbMarkIn1Arr;       ///< Per-row flag: @c true = this row has a difference to mark in file 1.
	std::vector<bool>    m_pbMarkIn2Arr;       ///< Per-row flag: @c true = this row has a difference to mark in file 2.
	ComparisonMatrix     m_matrix;             ///< Result comparison matrix (match percentages per column pair).
	std::vector<bool>    m_pbGreenClms1;       ///< Per-column flag: @c true = column fully matched in table 1.
	std::vector<bool>    m_pbGreenClms2;       ///< Per-column flag: @c true = column fully matched in table 2.
	std::vector<long>    m_pnFoundDifferences; ///< List of Excel row indices where differences were found.

	// Cross-table similarity results
	std::vector<SimilaritiesAcrossTables> m_vecSimilaritiesAcrossTables;       ///< Raw similarity scores (one entry per column pair).
	std::vector<SimilaritiesAcrossTables> m_vecSimilaritiesAcrossTablesSorted; ///< Same as above, sorted by descending score.
	long m_nSelectedDifference; ///< Index of the difference currently selected in the differences combo.

	CMFCRibbonBar* m_pRibbon;  ///< Pointer to the application ribbon bar.

	Table m_Table1; ///< Descriptor for the first compared table (dimensions, header/data rows).
	Table m_Table2; ///< Descriptor for the second compared table.

	ExcelConnector m_excel1;   ///< OLE connection to the first Excel workbook.
	ExcelConnector m_excel2;   ///< OLE connection to the second Excel workbook.
	CApplication   m_App;      ///< Shared Excel Application COM object.
	CString        m_szFilename1; ///< Full path of the first file.
	CString        m_szFilename2; ///< Full path of the second file.
	std::map<CString, long> m_mapTmpMap1; ///< Temporary key→row map for table 1 (used by findSims).
	std::map<CString, long> m_mapTmpMap2; ///< Temporary key→row map for table 2 (used by findSims).

	// UI state
	int   m_nUiToBeRefreshed; ///< Bitmask of UI areas that need refreshing on the next idle cycle.
	float m_fZoom;            ///< Current zoom factor for the comparison matrix.

	// Ribbon UI element pointers
	CMFCRibbonProgressBar* m_pProgressBar1;    ///< Ribbon progress bar for comparison pass 1.
	CMFCRibbonProgressBar* m_pProgressBar2;    ///< Ribbon progress bar for comparison pass 2.
	CMFCRibbonProgressBar* m_pKeyProgressBar1; ///< Ribbon progress bar for key-search pass 1.
	CMFCRibbonProgressBar* m_pKeyProgressBar2; ///< Ribbon progress bar for key-search pass 2.
	CMFCRibbonComboBox*    m_pCombo2;          ///< Key-column combo box for table 2.
	CMFCRibbonComboBox*    m_pSheetCombo1;     ///< Sheet-selection combo box for file 1.
	CMFCRibbonComboBox*    m_pSheetCombo2;     ///< Sheet-selection combo box for file 2.
	CMFCRibbonEdit*        m_pSpinner1_Fdata;  ///< "First data row" spinner for table 1.
	CMFCRibbonEdit*        m_pSpinner1_Names;  ///< "Header row" spinner for table 1.
	CMFCRibbonEdit*        m_pSpinner2_Fdata;  ///< "First data row" spinner for table 2.
	CMFCRibbonEdit*        m_pSpinner2_Names;  ///< "Header row" spinner for table 2.
	CMFCRibbonCheckBox*    m_pMarkIn1;         ///< "Mark in file 1" checkbox.
	CMFCRibbonCheckBox*    m_pMarkIn2;         ///< "Mark in file 2" checkbox.
	CMFCRibbonSlider*      m_pSlider;          ///< Similarity-threshold slider.
	CMFCRibbonButton*      m_pUnhideExcel;     ///< "Bring Excel to front" button.
	CMFCRibbonCheckBox*    m_pVerifyKeys;      ///< "Verify keys" checkbox.
	CMFCRibbonCheckBox*    m_pSameNames;       ///< "Same column names only" checkbox.
	CMFCRibbonColorButton* m_pColorPicker1;    ///< Colour picker for highlight colour 1.
	CMFCRibbonColorButton* m_pColorPicker2;    ///< Colour picker for highlight colour 2.
	CMFCRibbonCheckBox*    m_pAuto;            ///< "Auto-mark" checkbox.
	CMFCRibbonComboBox*    m_pFoundDifferences; ///< Differences navigation combo box.
	CMFCRibbonLabel*       m_pLabel0;          ///< Info label 0 (file/sheet name display).
	CMFCRibbonLabel*       m_pLabel1;          ///< Info label 1.
	CMFCRibbonLabel*       m_pLabel2;          ///< Info label 2.
	CMFCRibbonCheckBox*    m_pToFront;         ///< "Bring to front" checkbox.
	CMFCRibbonCheckBox*    m_pShowSims;        ///< "Show similar columns" checkbox.
	CMFCRibbonButton*      m_pCreateNewKeys;   ///< "Suggest keys" button.
	CMFCRibbonButton*      m_pButton2;         ///< Secondary action button.
	CMFCRibbonCheckBox*    m_pUseIndices;      ///< "Use key indexing" checkbox.
	CMFCRibbonEdit*        m_pRows1;           ///< Row-count edit for table 1.
	CMFCRibbonEdit*        m_pCols1;           ///< Column-count edit for table 1.
	CMFCRibbonEdit*        m_pRows2;           ///< Row-count edit for table 2.
	CMFCRibbonEdit*        m_pCols2;           ///< Column-count edit for table 2.

	// Scroll / view state
	bool       m_bToFront;              ///< If @c true, bring Excel to the foreground after navigation.
	int        m_nScrolled_X;           ///< Current horizontal scroll position in pixels.
	int        m_nScrolled_Y;           ///< Current vertical scroll position in pixels.
	ChosenCell M_CCell;                 ///< Currently hovered matrix cell.
	ChosenCell m_CClickedCell;          ///< Last double-clicked cell.
	ChosenCell m_CPrevClickedCell;      ///< Previously double-clicked cell (for toggle behaviour).
	ChosenCell m_OldCell;               ///< Cell hovered in the previous mouse-move event.
	VisTopLeft m_VisTopLeft;            ///< Top-left visible cell index (after scroll).
	bool m_bIn1file;                    ///< @c true when navigating to differences in file 1 is active.
	bool m_bIn2file;                    ///< @c true when navigating to differences in file 2 is active.
	bool m_bToDisplaySimilarClms;       ///< @c true when the similarity-line overlay is shown.
	bool m_bXSimilarityComputed;        ///< @c true after the column-similarity pass has completed.
	bool m_bAutoMark;                   ///< Mirror of the auto-mark checkbox state.
	bool m_bVerifyKeys;                 ///< Mirror of the "Verify keys" checkbox state.
	bool m_bToInitSB;                   ///< @c true when scroll bars need to be reinitialised.
	int  m_nCellWidth;                  ///< Width of a single matrix cell in pixels.
	int  m_nCellHeight;                 ///< Height of a single matrix cell in pixels.
	int  m_nRibbonWidth;                ///< Cached ribbon width (subtracted from client width).
	int  m_nViewWidth;                  ///< Width of the visible matrix area in pixels.
	int  m_nViewHeight;                 ///< Height of the visible matrix area in pixels.
	int  m_nHScrollPos;                 ///< Current horizontal scroll-bar position (cell units).
	int  m_nVScrollPos;                 ///< Current vertical scroll-bar position (cell units).
	int  m_nHPageSize;                  ///< Horizontal page size for the scroll bar.
	int  m_nVPageSize;                  ///< Vertical page size for the scroll bar.
	bool m_bOnlyPcnt;                   ///< If @c true, cells show only a percentage (no colour).
	bool m_bForceNotOnlyPcnt;           ///< Overrides @c m_bOnlyPcnt when set.
	int  m_nSldr;                       ///< Cached similarity-threshold slider value (0–100).
	CPen m_SimsPens[256];               ///< Array of pens for drawing similarity lines (one per colour step).
	CPen m_KeyCurvePen;                 ///< Pen used to draw key-pair curves in column headers.
	bool m_bUseIndexes;                 ///< @c true when index-based key disambiguation is active.
	bool m_bNewFile1;                   ///< @c true if file 1 was replaced since the last comparison.
	bool m_bNewFile2;                   ///< @c true if file 2 was replaced since the last comparison.
	CString m_szRsrcs;                  ///< Path to the resource satellite DLL (if used).
};

