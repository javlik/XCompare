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

// MainFrm.h : interface of the CMainFrame class
//

#pragma once
#include "ChildView.h"

/**
 * @brief The application's main frame window, hosting the ribbon bar, status bar and child view.
 *
 * Derives from @c CFrameWndEx. Owns the @c CMFCRibbonBar, @c CMFCRibbonStatusBar and
 * the @c CChildView. Routes most commands to the child view via @c OnCmdMsg.
 */
class CMainFrame : public CFrameWndEx
{

public:
    CMainFrame();

protected:
    DECLARE_DYNAMIC(CMainFrame)

    // Attributes
public:
    // Operations
public:
    // Overrides
public:
    /** @brief Adjusts the window class before creation (removes the client-edge style). */
    virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
    /** @brief Forwards commands to the child view before falling back to the base class. */
    virtual BOOL OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo);

    // Implementation
public:
    virtual ~CMainFrame();
#ifdef _DEBUG
    virtual void AssertValid() const;
    virtual void Dump(CDumpContext& dc) const;
#endif

protected: // control bar embedded members
    CMFCRibbonBar m_wndRibbonBar;
    CMFCRibbonApplicationButton m_MainButton;
    CMFCToolBarImages m_PanelImages;
    CMFCRibbonStatusBar m_wndStatusBar;
    CChildView m_wndView;

    // Generated message map functions
protected:
    /** @brief Creates the child view, ribbon bar and status bar. */
    afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
    /** @brief Forwards keyboard focus to the child view. */
    afx_msg void OnSetFocus(CWnd* pOldWnd);
    /** @brief Applies the selected visual theme to the ribbon and redraws the frame. */
    afx_msg void OnApplicationLook(UINT id);
    /** @brief Checks/unchecks the currently active visual-theme menu item. */
    afx_msg void OnUpdateApplicationLook(CCmdUI* pCmdUI);
    DECLARE_MESSAGE_MAP()

public:
    /** @brief Displays @p text in the ribbon status bar's first pane. */
    void updateStatusBar(CString text);
    /** @brief Translates a horizontal mouse-wheel event into a horizontal scroll on the child view. */
    afx_msg void OnMouseHWheel(UINT nFlags, short zDelta, CPoint pt);
};
