// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface 
// (the "Fluent UI") and is provided only as referential material to supplement the 
// Microsoft Foundation Classes Reference and related electronic documentation 
// included with the MFC C++ library software.  
// License terms to copy, use or distribute the Fluent UI are available separately.  
// To learn more about our Fluent UI licensing program, please visit 
// http://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// ChildView.h : interface of the CChildView class
//


#pragma once
#include "Constants.h"
#include "TableData.h"
#include "ComparisonMatrix.h"
#include "ExcelConnector.h"
#include "ComparisonEngine.h"
#include <vector>
#include <map>

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
public:
	afx_msg void OnPickFirstFile();
	afx_msg void OnPickSecondFile();
	afx_msg void OnCreateMatrix();
	afx_msg void OnUpdatePickFirstSheet(CCmdUI *pCmdUI);
	afx_msg void OnUpdateCreateMatrix(CCmdUI *pCmdUI);
//	afx_msg void OnMouseHWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdateFilename1(CCmdUI *pCmdUI);
	afx_msg void OnUpdateFilename2(CCmdUI *pCmdUI);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdatePickSecondSheet(CCmdUI *pCmdUI);
	afx_msg void OnUpdateProgress1(CCmdUI *pCmdUI);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);

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
	int mxGet(int x, int y);
	void mxClear(int x, int y);
	int mxPut(int x, int y);
	bool mxMarkedGet(int x, int y);
	CString getCellValue1(int column, int row);
	CString getCellValue2(int column, int row);
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
public:
	//afx_msg void OnProgress2();
	void markInFiles();
	afx_msg void OnButton5();
	afx_msg void OnUpdateButton5(CCmdUI *pCmdUI);
	afx_msg void OnButton3();
	afx_msg void OnUpdateButton3(CCmdUI *pCmdUI);
	afx_msg void OnCheck3();
	afx_msg void OnUpdateCheck3(CCmdUI *pCmdUI);
	void makePrereq1(); // delegates to m_engine
	void makePrereq2(); // delegates to m_engine
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
	
	

	int createTempKeyArrays1();
	int createTempKeyArrays2();
	void clearPossibleKeys();
	void sort3(int & a, int & b, int & c);
public:
	bool mutualCheck();
	int checkKeys(int tab1);
	int deleteKey(int table, int column);
	void setKey(int table, int column);
	void deleteAllKeys();
	bool areThereAnyKeys();
	bool isThisAKey(int table, int column);
	int getNthKey(int table, int key);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	void setNthKey(int n, int col1, int col2);
	void insertKeyAt(int n, int col1, int col2); // kept for compatibility
	void deleteKeyAt(int n);
	void pushKey(int col1, int col2);
	bool usePossibleKeys();
	int getNumberOfPossibleKeys();
	void sortExaminedKeys(int table);
	int sumExaminedKeys(int table, int nmax);
	bool is2BExaminedOnce(int table, int max);
	bool getSimilarKeyProbability(int table, int max);
	int getNthEntropy(int table, int n);
	int CalculateEntropyRank(int table);
	bool isEntropyStored(int table, int clm, int max);
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
	BOOL         m_bUniqueKeys1; // keys in table 1 are unique
	BOOL         m_bUniqueKeys2; // keys in table 2 are unique
	bool         m_bLockPrg1;    // thread 1 is running
	bool         m_bLockPrg2;    // thread 2 is running
	// m_NotUniqueKeys1/2 are now in m_engine.m_NotUniqueKeys1/2
	ComparisonEngine m_engine;

private:
	// Synchronisation between threads
	bool m_bWaitingForKeys;
	bool m_bKeys1done;
	bool m_bKeys2done;
	bool m_bKeysGathering1done;
	bool m_bKeysGathering2done;

	// Algorithm parameters
	int     m_nComplexity;
	CString m_szRsltTxt;

	// Large checked-key buffers (one entry per attempted key combination)
	std::vector<unsigned long long> m_nCheckedKeys1;
	std::vector<unsigned long long> m_nCheckedKeys2;
	int m_nCheckedKeysCounter1;
	int m_nCheckedKeysCounter2;

	// Colour palette (20 user-selectable highlight colours)
	Palette m_Palette[20];

	// Key-suggestion data structures
	PossibleKeys m_PossibleKeys1[256];
	PossibleKeys m_PossibleKeys2[256];
	KeyPair      m_KeyPair[256];
	int          m_nKeyPairCounter;
	BestKeyComb  m_BestKeyComb;
	int          m_nPossibleKeyCounter1;
	int          m_nPossibleKeyCounter2;
	long         m_nInvEntropy1[256];
	long         m_nInvEntropy2[256];
	int          m_nSortedEntropy1[256];
	int          m_nSortedEntropy2[256];

	// Prerequisite validity is now tracked inside m_engine
	int  m_nOldx;
	int  m_nOldy;
	int  m_nChosenColor1;
	int  m_nChosenColor2;
	Clnt m_Clnt;
	bool m_bDoAutoMark;
	int  m_nNatrixDone;
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

	// Entropy tracking for key suggestion
	int m_nExaminedKeys1[SUGKEYS + 4];
	int m_nExaminedKeys2[SUGKEYS + 4];
	int m_nTmpKeys1[SUGKEYS + 4];
	int m_nTmpKeys2[SUGKEYS + 4];

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
	std::map<CString, long> m_mapTmpMap1;
	std::map<CString, long> m_mapTmpMap2;

	// m_engine is now public (accessed by thread procs)

	// UI state
	int   m_nUiToBeRefreshed;
	float m_fZoom;
	int   m_nPrgval1;

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

