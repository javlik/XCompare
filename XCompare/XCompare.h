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

// XCompare.h : main header file for the XCompare application
//
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"       // main symbols


/**
 * @brief The application singleton class for XCompare.
 *
 * Derives from @c CWinAppEx and manages application-level initialisation,
 * OLE startup, the main frame window, and the ribbon/look persistence.
 */
class CXCompareApp : public CWinAppEx
{
public:
	CXCompareApp();

// Overrides
public:
	/** @brief Initialises OLE, creates the main frame window, and shows it. */
	virtual BOOL InitInstance();
	/** @brief Terminates OLE and cleans up before the process exits. */
	virtual int ExitInstance();

// Implementation

public:
	UINT  m_nAppLook; ///< ID of the currently active visual style (persisted in the registry).
	/** @brief Registers the Edit pop-up context menu before the state is loaded. */
	virtual void PreLoadState();
	/** @brief Placeholder for loading custom persistent state (currently empty). */
	virtual void LoadCustomState();
	/** @brief Placeholder for saving custom persistent state (currently empty). */
	virtual void SaveCustomState();

	/** @brief Displays the About dialog box. */
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CXCompareApp theApp;
