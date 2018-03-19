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
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "CCellFormat.h"
#include "Cnterior.h"

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
	afx_msg void OnUpdateSpin2(CCmdUI *pCmdUI);
	afx_msg void OnUpdateSpin3(CCmdUI *pCmdUI);
	afx_msg void OnUpdateProgress1(CCmdUI *pCmdUI);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	CWorksheets GetWorksheets(CString TempBookName);
	afx_msg
		CWorksheets GetWorksheets1(CString TempBookName);
	CWorksheets GetWorksheets2(CString TempBookName);
	void OnPickFirstSheet();
	afx_msg void OnSpin3();
	afx_msg void OnSpin1Names();
	afx_msg void OnUpdateSpin1Names(CCmdUI *pCmdUI);
	afx_msg void OnUpdateSpin1Fdata(CCmdUI *pCmdUI);
	afx_msg void OnSpin1Fdata();
	afx_msg void OnUpdateKey11(CCmdUI *pCmdUI);
	afx_msg void OnUpdateKey12(CCmdUI *pCmdUI);
	afx_msg void OnUpdateKey13(CCmdUI *pCmdUI);
	void updateCombos1();
	afx_msg void OnKey11();
	afx_msg void OnKey12();
	afx_msg void OnKey13();
	afx_msg void OnPickSecondSheet();
	void updateCombos2();
	afx_msg void OnUpdateSpin2Fdata(CCmdUI *pCmdUI);
	afx_msg void OnSpin2Fdata();
	afx_msg void OnUpdateSpin2Names(CCmdUI *pCmdUI);
	afx_msg void OnSpin2Names();
	afx_msg void OnKey21();
	afx_msg void OnUpdateKey21(CCmdUI *pCmdUI);
	afx_msg void OnUpdateKey22(CCmdUI *pCmdUI);
	afx_msg void OnUpdateKey23(CCmdUI *pCmdUI);
	afx_msg void OnKey22();
	afx_msg void OnKey23();
	void makeCharArr1();
	void makeCharArr2();
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	void mxClear(int x, int y);
	int mxPut(int x, int y);
	int mxGet(int x, int y);
	bool mxMarkedGet(int x, int y);
	void checkEmptiness1();
	void checkEmptiness2();
	void checkEmptiness();
	bool checkKeysUniqueness1();
	bool checkKeysUniqueness2();
	//void firstPass(int thrdIdx);
	void firstPass();
	int createKeyArrays1();
	int createKeyArrays2();
	int createKeyArrays();
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
public:
	//afx_msg void OnProgress2();
	void markInFiles();
	afx_msg void OnButton5();
	afx_msg void OnUpdateButton5(CCmdUI *pCmdUI);
	afx_msg void OnButton3();
	afx_msg void OnUpdateButton3(CCmdUI *pCmdUI);
	afx_msg void OnCheck3();
	afx_msg void OnUpdateCheck3(CCmdUI *pCmdUI);
	void makePrereq1();
	void makePrereq2();
	afx_msg void OnUpdateProgress3(CCmdUI *pCmdUI);
	void resolveAutoMark();
	void DrainMsgQueue(void);
	afx_msg void OnDiffslist();
	afx_msg void OnUpdateDiffslist(CCmdUI *pCmdUI);
	afx_msg void OnPckd0();
	afx_msg void OnPckd1();
	afx_msg void OnPckd2();
	afx_msg void OnUpdatePckd0(CCmdUI *pCmdUI);
	afx_msg void OnUpdatePckd1(CCmdUI *pCmdUI);
	afx_msg void OnUpdatePckd2(CCmdUI *pCmdUI);
	afx_msg void OnSelIn1();
	afx_msg void OnSelIn2();
	afx_msg void OnInSel1();
	afx_msg void OnInSel2();
	afx_msg void OnSel1();
	int rowFromCombo();
	afx_msg void OnButton6();
	afx_msg void OnPut2front();
	afx_msg void OnUpdatePut2front(CCmdUI *pCmdUI);
	void decrementThreadCnt();
};

