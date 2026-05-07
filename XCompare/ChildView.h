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

// CChildView window

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
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);

// Implementation
public:
	virtual ~CChildView();

	// Generated message map functions
protected:
	afx_msg void OnPaint();
	DECLARE_MESSAGE_MAP()

	// OnPaint helper context: groups all temporary GDI objects and bounds
	struct PaintCtx
	{
		CPen   *pen1, *pen2, *pen3, *pen4, *pen5, *pen6,
		       *pen7, *pen8, *pen9, *pen10, *pen11, *pen12;
		CBrush *brush0, *brush1, *brush2, *brush3,
		       *brush4, *brush5, *brush6, *brush7;
		CFont  *font1, *font2, *font3, *font4,
		       *font1B, *font2B, *font1C, *font2C;
		int    bnd_X_min, bnd_X_max, bnd_Y_min, bnd_Y_max;
	};

	void paintInfoArea      (CDC& dc, PaintCtx& ctx);
	void paintGridLines     (CDC& dc, PaintCtx& ctx);
	void paintRowHeaders    (CDC& dc, PaintCtx& ctx);
	void paintColumnHeaders (CDC& dc, PaintCtx& ctx);
	void paintMatrixCells   (CDC& dc, PaintCtx& ctx);
	void paintSimilarityLines(CDC& dc, PaintCtx& ctx);

public:
	afx_msg void OnPickFirstFile();
	afx_msg void OnPickSecondFile();
    void pickFile(bool* pNewFile, ExcelConnector* pExcel, Table* pTable, CMFCRibbonComboBox* pSheetCombo, CString* pFilename);
	afx_msg void OnCreateMatrix();
	afx_msg void OnUpdatePickFirstSheet(CCmdUI *pCmdUI);
	afx_msg void OnUpdateCreateMatrix(CCmdUI *pCmdUI);
//	afx_msg void OnMouseHWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdateFilename1(CCmdUI *pCmdUI);
	afx_msg void OnUpdateFilename2(CCmdUI *pCmdUI);
	void updateFileName(CCmdUI* pCmdUI, CString* pszFilename, int idString);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdatePickSecondSheet(CCmdUI *pCmdUI);
	afx_msg void OnUpdateProgress1(CCmdUI *pCmdUI);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);

	void pickSheet(ExcelConnector* pExcel, Table* pTable, CMFCRibbonComboBox* pSheetCombo, CMFCRibbonEdit* pSpinner_Names, CMFCRibbonEdit* pSpinner_Fdata, CMFCRibbonEdit* pRows, CMFCRibbonEdit* pCols);
	void OnPickFirstSheet();
	afx_msg void OnSpin1Names();
	afx_msg void OnUpdateSpin1Names(CCmdUI *pCmdUI);
	afx_msg void OnUpdateSpin1Fdata(CCmdUI *pCmdUI);
	afx_msg void OnSpin1Fdata();
	void updateCombos1();
	afx_msg void OnPickSecondSheet();
	void updateCombos2();
	afx_msg void OnUpdateSpin2Fdata(CCmdUI *pCmdUI);
	afx_msg void OnSpin2Fdata();
	afx_msg void OnUpdateSpin2Names(CCmdUI *pCmdUI);
	afx_msg void OnSpin2Names();
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	void firstPass();
	int createKeyArrays1();
	int createKeyArrays2();
	bool checkKeysUniqueness1();
	bool checkKeysUniqueness2();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnSlider2();
	afx_msg void OnUpdateSlider2(CCmdUI *pCmdUI);
	afx_msg void OnCheck4();
	afx_msg void OnUpdateCheck4(CCmdUI *pCmdUI);
	afx_msg void OnCheck5();
	afx_msg void OnUpdateCheck5(CCmdUI *pCmdUI);
	afx_msg void OnButton2();
//	CString convertR1C1();
	CString convertR1C1(int row, int clm);
	void markIn1(int row, int clm);
	void markIn2(int row, int clm);
	void initScrollBars();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnUpdateProgress2(CCmdUI *pCmdUI);

	afx_msg void OnUpdateCheck2(CCmdUI *pCmdUI);
	afx_msg void OnCheck2();
//	UINT JobThread();
	afx_msg void OnUpdateButton2(CCmdUI *pCmdUI);


	afx_msg void OnCheck7();
	afx_msg void OnUpdateCheck7(CCmdUI *pCmdUI);
protected:
	afx_msg LRESULT OnCmUpdateProgress(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmUpdateProgress2(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmUpdateProgress3(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmUpdateKeyProgress1(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmUpdateKeyProgress2(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmKeys1Done(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmKeys2Done(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmGathering1Done(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmGathering2Done(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmKeysFound(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmKeysNotFound(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmMarkingReady(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmSims1Done(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmSims2Done(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmFirstPassDone(WPARAM wParam, LPARAM lParam);
public:
	//afx_msg void OnProgress2();
	void markInFiles();
	afx_msg void OnButton5();
	afx_msg void OnUpdateButton5(CCmdUI *pCmdUI);
	afx_msg void OnButton3();
	afx_msg void OnUpdateButton3(CCmdUI *pCmdUI);
	afx_msg void OnCheck3();
	afx_msg void OnUpdateCheck3(CCmdUI *pCmdUI);
	void resolveAutoMark();
	void DrainMsgQueue(void);
	afx_msg void OnDiffslist();
	afx_msg void OnUpdateDiffslist(CCmdUI *pCmdUI);
	afx_msg void OnSel1();
	int rowFromCombo();
	afx_msg void OnButton6();
	afx_msg void OnPut2front();
	afx_msg void OnUpdatePut2front(CCmdUI *pCmdUI);
	void suggestKeys1();
	void suggestKeys2();
private:
	//void suggestKeys();
	
	

	int createTempKeyArrays1(); // kept as shell (body moved to KeyFinder)
	int createTempKeyArrays2(); // kept as shell (body moved to KeyFinder)
	void clearPossibleKeys();
	void sort3(int & a, int & b, int & c);
	void findSimsRange(int c_i1_start, int c_i1_end, UINT progressMsg, UINT doneMsg, bool useTmp);
public:
	bool mutualCheck();
	void deleteAllKeys();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	bool usePossibleKeys();
	int getNumberOfPossibleKeys();
	// sortExaminedKeys, sumExaminedKeys, is2BExaminedOnce, getSimilarKeyProbability,
	// getNthEntropy, CalculateEntropyRank, isEntropyStored -- moved to KeyFinder (private)
	afx_msg void OnUpdateCombo2(CCmdUI *pCmdUI);
	afx_msg void OnCombo2();
	int getNumberOfPossibleKeys(int table, int order, int item);
	void findSims();
	void findSims1();
	void findSims2();
	afx_msg void OnSimilarpaircheckbox();
	afx_msg void OnUpdateSimilarpaircheckbox(CCmdUI *pCmdUI);
	afx_msg void OnFindrelBtn();
	afx_msg void OnIdxcrtBtn();
	afx_msg void OnUpdateKeyProgress1(CCmdUI *pCmdUI);
	afx_msg void OnUpdateKeyProgress2(CCmdUI *pCmdUI);
	void finishFindRelations();
	afx_msg void OnUpdateIdxCheckbox(CCmdUI *pCmdUI);
	int ReverseFind(LPCTSTR lpszData, LPCTSTR lpszSub, int startpos);
	afx_msg void OnCheckIdx();
	afx_msg void OnUpdateCheckIdx(CCmdUI *pCmdUI);
	afx_msg void OnUsidxCheck();
	afx_msg void OnUpdateUsidxCheck(CCmdUI *pCmdUI);
	afx_msg void OnUpdateRows1(CCmdUI *pCmdUI);
	afx_msg void OnRows1();
	afx_msg void OnUpdateCols1(CCmdUI *pCmdUI);
	afx_msg void OnCols1();
	afx_msg void OnUpdateRows2(CCmdUI *pCmdUI);
	afx_msg void OnRows2();
	afx_msg void OnCols2();
	afx_msg void OnUpdateCols2(CCmdUI *pCmdUI);

	// ---------------------------------------------------------------
	// Data members (moved from .cpp file-scope globals)
	// ---------------------------------------------------------------

	// Public: accessed directly by worker thread functions
public:
	std::atomic<bool>         m_bUniqueKeys1{false}; // keys in table 1 are unique
	std::atomic<bool>         m_bUniqueKeys2{false}; // keys in table 2 are unique
	std::atomic<bool>         m_bLockPrg1{false};    // thread 1 is running
	std::atomic<bool>         m_bLockPrg2{false};    // thread 2 is running
	// m_NotUniqueKeys1/2 are now in m_engine.m_NotUniqueKeys1/2
	ComparisonEngine m_engine;
	KeyFinder        m_keyFinder;

private:
	// Synchronisation between threads
	std::atomic<bool> m_bWaitingForKeys{false};
	std::atomic<bool> m_bKeys1done{false};
	std::atomic<bool> m_bKeys2done{false};
	std::atomic<bool> m_bKeysGathering1done{false};
	std::atomic<bool> m_bKeysGathering2done{false};

	// Algorithm parameters
	CString m_szRsltTxt;

	// Colour palette (20 user-selectable highlight colours)
	Palette m_Palette[20];

	// Prerequisite validity is now tracked inside m_engine
	int  m_nOldx;
	int  m_nOldy;
	int  m_nChosenColor1;
	int  m_nChosenColor2;
	Clnt m_Clnt;
	bool m_bDoAutoMark;
	int  m_nMatrixDone;
	int  m_nPrereqDone;
	bool m_bMarkIdentCols;
	bool m_bSameNames;
	int  m_nEffMax;

	// Dynamic data arrays
	std::vector<bool>    m_pbMarkIn1Arr;    // cells to mark in file 1
	std::vector<bool>    m_pbMarkIn2Arr;    // cells to mark in file 2
	ComparisonMatrix     m_matrix;          // result comparison matrix
	std::vector<bool>    m_pbGreenClms1;    // fully-matched column flags, table 1
	std::vector<bool>    m_pbGreenClms2;    // fully-matched column flags, table 2
	std::vector<long>    m_pnFoundDifferences; // difference row indices per column
	// Arrays now owned by m_engine: m_pchMainArr1/2, m_pszKeyArr11/21, m_pbKeyMissing1/2, m_pbEmptyClms1/2
	// Entropy tracking (now inside m_keyFinder): m_nExaminedKeys1/2, m_nInvEntropy1/2, m_nSortedEntropy1/2

	// Cross-table similarity results
	std::vector<SimilaritiesAcrossTables> m_vecSimilaritiesAcrossTables;
	std::vector<SimilaritiesAcrossTables> m_vecSimilaritiesAcrossTablesSorted;
	long m_nSelectedDifference;

	// MFC Ribbon pointer
	CMFCRibbonBar* m_pRibbon;

	// Table descriptors
	Table m_Table1;
	Table m_Table2;

	// Excel connections (one per compared table) + shared application object
	ExcelConnector m_excel1;
	ExcelConnector m_excel2;
	CApplication   m_App;
	CString        m_szFilename1;
	CString        m_szFilename2;
	// m_Map1/m_Map2 are now inside m_engine
	// m_mapTmpMap1/2 are now inside m_keyFinder; CChildView also keeps its own copies for findSims
	std::map<CString, long> m_mapTmpMap1;
	std::map<CString, long> m_mapTmpMap2;

	// m_engine and m_keyFinder are now public (accessed by thread procs)

	// UI state
	int   m_nUiToBeRefreshed;
	float m_fZoom;

	// Ribbon UI element pointers
	CMFCRibbonProgressBar* m_pProgressBar1;
	CMFCRibbonProgressBar* m_pProgressBar2;
	CMFCRibbonProgressBar* m_pKeyProgressBar1;
	CMFCRibbonProgressBar* m_pKeyProgressBar2;
	CMFCRibbonComboBox*    m_pCombo2;
	CMFCRibbonComboBox*    m_pSheetCombo1;
	CMFCRibbonComboBox*    m_pSheetCombo2;
	CMFCRibbonEdit*        m_pSpinner1_Fdata;
	CMFCRibbonEdit*        m_pSpinner1_Names;
	CMFCRibbonEdit*        m_pSpinner2_Fdata;
	CMFCRibbonEdit*        m_pSpinner2_Names;
	CMFCRibbonCheckBox*    m_pMarkIn1;
	CMFCRibbonCheckBox*    m_pMarkIn2;
	CMFCRibbonSlider*      m_pSlider;
	CMFCRibbonButton*      m_pUnhideExcel;
	CMFCRibbonCheckBox*    m_pVerifyKeys;
	CMFCRibbonCheckBox*    m_pSameNames;
	CMFCRibbonColorButton* m_pColorPicker1;
	CMFCRibbonColorButton* m_pColorPicker2;
	CMFCRibbonCheckBox*    m_pAuto;
	CMFCRibbonComboBox*    m_pFoundDifferences;
	CMFCRibbonLabel*       m_pLabel0;
	CMFCRibbonLabel*       m_pLabel1;
	CMFCRibbonLabel*       m_pLabel2;
	CMFCRibbonCheckBox*    m_pToFront;
	CMFCRibbonCheckBox*    m_pShowSims;
	CMFCRibbonButton*      m_pCreateNewKeys;
	CMFCRibbonButton*      m_pButton2;
	CMFCRibbonCheckBox*    m_pUseIndices;
	CMFCRibbonEdit*        m_pRows1;
	CMFCRibbonEdit*        m_pCols1;
	CMFCRibbonEdit*        m_pRows2;
	CMFCRibbonEdit*        m_pCols2;

	// Scroll / view state
	bool       m_bToFront;
	int        m_nScrolled_X;
	int        m_nScrolled_Y;
	ChosenCell M_CCell;
	ChosenCell m_CClickedCell;
	ChosenCell m_CPrevClickedCell;
	ChosenCell m_OldCell;
	VisTopLeft m_VisTopLeft;
	bool m_bIn1file;
	bool m_bIn2file;
	bool m_bToDisplaySimilarClms;
	bool m_bXSimilarityComputed;
	bool m_bAutoMark;
	bool m_bVerifyKeys;
	bool m_bToInitSB;
	int  m_nCellWidth;
	int  m_nCellHeight;
	int  m_nRibbonWidth;
	int  m_nViewWidth;
	int  m_nViewHeight;
	int  m_nHScrollPos;
	int  m_nVScrollPos;
	int  m_nHPageSize;
	int  m_nVPageSize;
	bool m_bOnlyPcnt;
	bool m_bForceNotOnlyPcnt;
	int  m_nSldr;
	CPen m_SimsPens[256];
	CPen m_KeyCurvePen;
	bool m_bUseIndexes;
	bool m_bNewFile1;
	bool m_bNewFile2;
	CString m_szRsrcs;
};

