#include "stdafx.h"
#include <cstring>
#include <map>
#include <vector>
#include "XCompare.h"
#include "ChildView.h"
#include "MainFrm.h"
#include "Msg.h"
extern CMainFrame* g_pMainFrame; // pointer to FrameWindow
#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// Forward declarations of worker thread functions
UINT MyThreadProc(LPVOID pParam);
UINT MyThreadProc2(LPVOID pParam);
UINT MyThreadProc3(LPVOID pParam);
UINT CreateKeys1ThreadProc(LPVOID pParam);
UINT CreateKeys2ThreadProc(LPVOID pParam);
UINT makePrereq1ThreadProc(LPVOID pParam);
UINT makePrereq2ThreadProc(LPVOID pParam);
UINT SuggestKeys1ThreadProc(LPVOID pParam);
UINT SuggestKeys2ThreadProc(LPVOID pParam);
UINT MutualCheckThreadProc(LPVOID pParam);
UINT FindSimsThreadProc(LPVOID pParam);
UINT FindSimsThreadProc1(LPVOID pParam);
UINT FindSimsThreadProc2(LPVOID pParam);
CChildView::CChildView()
{
	//threadCnt = 1;
	m_szRsrcs = L"";
	m_bToFront = false;
	m_nSelectedDifference = 0;
	m_bForceNotOnlyPcnt = true;
	// m_bPrereq1valid / m_bPrereq2valid are now inside m_engine
	m_nChosenColor1 = 13;
	m_nChosenColor2 = 13;
	m_Palette[0] = { 0,   0,   0 };
	m_Palette[1] = { 128,   0,   0 };
	m_Palette[2] = { 0,   128,   0 };
	m_Palette[3] = { 128, 128,   0 };
	m_Palette[4] = { 0,   0,   128 };
	m_Palette[5] = { 128,   0, 128 };
	m_Palette[6] = { 0,   128, 128 };
	m_Palette[7] = { 192, 192, 192 };
	m_Palette[8] = { 192, 220, 192 };
	m_Palette[9] = { 166, 202, 240 };
	m_Palette[10] = { 255, 251, 240 };
	m_Palette[11] = { 160, 160, 164 };
	m_Palette[12] = { 128, 128, 128 };
	m_Palette[13] = { 255,   0,   0 };
	m_Palette[14] = { 0,   255,   0 };
	m_Palette[15] = { 255, 255,   0 };
	m_Palette[16] = { 0,   0,   255 };
	m_Palette[17] = { 255,   0, 255 };
	m_Palette[18] = { 0, 255, 255 };
	m_Palette[19] = { 255, 255, 255 };
	m_bToDisplaySimilarClms = false;
	m_bXSimilarityComputed = false;
	m_bAutoMark = false;
	m_bDoAutoMark = false;
	m_szRsltTxt = "";
	m_bUniqueKeys1 = false;
	m_bUniqueKeys2 = false;
	m_bLockPrg1 = false;
	m_bLockPrg2 = false;
	m_bVerifyKeys = false;
	m_szFilename1 = "";
	m_szFilename2 = "";
	m_nUiToBeRefreshed = 3;
	m_fZoom = 100;
	m_nPrgval1 = 100; // just for test
				   //CMFCRibbonBar* pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar(); // in the constructor is too early
	// ExcelConnector instances (m_excel1, m_excel2) initialise themselves in their own constructors
	// m_nCheckedKeys1/2, entropy and possible-keys state are now inside m_keyFinder
	// Members that were zero-initialised as globals but not yet in constructor
	m_bWaitingForKeys    = false;
	m_bKeys1done         = false;
	m_bKeys2done         = false;
	m_bKeysGathering1done = false;
	m_bKeysGathering2done = false;
	// m_nKeyPairCounter is now inside m_engine
	m_nOldx = 0; m_nOldy = 0;
	m_Clnt.w = 0; m_Clnt.h = 0;
	m_nNatrixDone  = 0;
	m_nPrereqDone  = 0;
	m_bMarkIdentCols = false;
	m_nCellWidth  = 0; m_nCellHeight  = 0;
	m_nRibbonWidth = 0;
	m_nViewWidth  = 0; m_nViewHeight  = 0;
	m_nHScrollPos = 0; m_nVScrollPos  = 0;
	m_nHPageSize  = 0; m_nVPageSize   = 0;
	m_nScrolled_X = 0; m_nScrolled_Y  = 0;
	m_CClickedCell.x = 0; m_CClickedCell.y = 0;
	m_CPrevClickedCell.x = 0; m_CPrevClickedCell.y = 0;
	m_pRibbon         = nullptr;
	m_pProgressBar1   = nullptr; m_pProgressBar2   = nullptr;
	m_pKeyProgressBar1 = nullptr; m_pKeyProgressBar2 = nullptr;
	m_pCombo2         = nullptr;
	m_pSheetCombo1    = nullptr; m_pSheetCombo2    = nullptr;
	m_pSpinner1_Fdata = nullptr; m_pSpinner1_Names = nullptr;
	m_pSpinner2_Fdata = nullptr; m_pSpinner2_Names = nullptr;
	m_pMarkIn1        = nullptr; m_pMarkIn2        = nullptr;
	m_pSlider         = nullptr; m_pUnhideExcel    = nullptr;
	m_pVerifyKeys     = nullptr; m_pSameNames      = nullptr;
	m_pColorPicker1   = nullptr; m_pColorPicker2   = nullptr;
	m_pAuto           = nullptr;
	m_pFoundDifferences = nullptr;
	m_pLabel0         = nullptr; m_pLabel1         = nullptr; m_pLabel2 = nullptr;
	m_pToFront        = nullptr; m_pShowSims       = nullptr;
	m_pCreateNewKeys  = nullptr; m_pButton2        = nullptr;
	m_pUseIndices     = nullptr;
	m_pRows1 = nullptr; m_pCols1 = nullptr;
	m_pRows2 = nullptr; m_pCols2 = nullptr;
	// entropy, possible-keys, examined-keys arrays are now inside m_keyFinder
	m_bIn1file = false;
	m_bIn2file = false;
	m_bSameNames = false;
	m_bOnlyPcnt = false;
	m_bToInitSB = true;
	m_VisTopLeft.left = 0;
	m_VisTopLeft.top = 0;
	// m_BestKeyComb is now inside m_keyFinder
	m_nSldr = 90;
	m_nEffMax = 0;
	M_CCell.x = 0; M_CCell.y = 0;
	m_OldCell.x = 0; m_OldCell.y = 0;
	m_Table1.NumberOfColumns = 0;
	m_Table2.NumberOfColumns = 0;
	for (int i = 0; i < 256; i++)
	{
		m_SimsPens[i].CreatePen(PS_ENDCAP_FLAT, (i) / 32 + 0.5,  RGB((255 - i) / 1.5 + 40, (255 - i) / 1.5 + 40, (255 - i) / 1.5 + 40));
	}
	m_KeyCurvePen.CreatePen(PS_ENDCAP_FLAT, 2, RGB(100, 150, 250));
	m_bUseIndexes = false;
	m_engine.m_bUseIndexes = false;
	m_bNewFile1 = false;
	m_bNewFile2 = false;
	// m_nComplexity is now inside m_keyFinder (default 100000)
}



/// <summary>
/// Finalizes an instance of the <see cref="CChildView"/> class.
/// </summary>
CChildView::~CChildView()
{
	// Dynamic arrays (std::vector) release memory automatically.
}



BEGIN_MESSAGE_MAP(CChildView, CWnd)
	ON_WM_PAINT()
	ON_COMMAND(ID_PICK_FIRST_FILE, &CChildView::OnPickFirstFile)
	ON_COMMAND(ID_PICK_SECOND_FILE, &CChildView::OnPickSecondFile)
	ON_COMMAND(ID_CREATE_MATRIX, &CChildView::OnCreateMatrix)
	ON_UPDATE_COMMAND_UI(ID_PICK_FIRST_SHEET, &CChildView::OnUpdatePickFirstSheet)
	ON_UPDATE_COMMAND_UI(ID_CREATE_MATRIX, &CChildView::OnUpdateCreateMatrix)
	//	ON_WM_MOUSEHWHEEL()
	ON_UPDATE_COMMAND_UI(IDC_FILENAME1, &CChildView::OnUpdateFilename1)
	ON_UPDATE_COMMAND_UI(IDC_FILENAME2, &CChildView::OnUpdateFilename2)
	ON_WM_MOUSEWHEEL()
	ON_UPDATE_COMMAND_UI(ID_PICK_SECOND_SHEET, &CChildView::OnUpdatePickSecondSheet)
	ON_UPDATE_COMMAND_UI(ID_PROGRESS1, &CChildView::OnUpdateProgress1)
	ON_WM_VSCROLL()
	ON_WM_HSCROLL()
	ON_COMMAND(ID_PICK_FIRST_SHEET, &CChildView::OnPickFirstSheet)
	ON_COMMAND(ID_SPIN1_NAMES, &CChildView::OnSpin1Names)
	ON_UPDATE_COMMAND_UI(ID_SPIN1_NAMES, &CChildView::OnUpdateSpin1Names)
	ON_UPDATE_COMMAND_UI(ID_SPIN1_FDATA, &CChildView::OnUpdateSpin1Fdata)
	ON_COMMAND(ID_SPIN1_FDATA, &CChildView::OnSpin1Fdata)
	ON_COMMAND(ID_PICK_SECOND_SHEET, &CChildView::OnPickSecondSheet)
	ON_UPDATE_COMMAND_UI(ID_SPIN2_FDATA, &CChildView::OnUpdateSpin2Fdata)
	ON_COMMAND(ID_SPIN2_FDATA, &CChildView::OnSpin2Fdata)
	ON_UPDATE_COMMAND_UI(ID_SPIN2_NAMES, &CChildView::OnUpdateSpin2Names)
	ON_COMMAND(ID_SPIN2_NAMES, &CChildView::OnSpin2Names)
	ON_WM_LBUTTONDBLCLK()
	ON_WM_MOUSEMOVE()
	ON_COMMAND(ID_SLIDER2, &CChildView::OnSlider2)
	ON_UPDATE_COMMAND_UI(ID_SLIDER2, &CChildView::OnUpdateSlider2)
	ON_COMMAND(ID_CHECK4, &CChildView::OnCheck4)
	ON_UPDATE_COMMAND_UI(ID_CHECK4, &CChildView::OnUpdateCheck4)
	ON_COMMAND(ID_CHECK5, &CChildView::OnCheck5)
	ON_UPDATE_COMMAND_UI(ID_CHECK5, &CChildView::OnUpdateCheck5)
	ON_COMMAND(ID_BUTTON2, &CChildView::OnButton2)
	ON_WM_SIZE()
	ON_WM_CREATE()
	ON_UPDATE_COMMAND_UI(ID_PROGRESS2, &CChildView::OnUpdateProgress1)
	ON_UPDATE_COMMAND_UI(ID_CHECK2, &CChildView::OnUpdateCheck2)
	ON_COMMAND(ID_CHECK2, &CChildView::OnCheck2)
	ON_UPDATE_COMMAND_UI(ID_BUTTON2, &CChildView::OnUpdateButton2)
	ON_COMMAND(ID_CHECK7, &CChildView::OnCheck7)
	ON_UPDATE_COMMAND_UI(ID_CHECK7, &CChildView::OnUpdateCheck7)
	ON_MESSAGE(CM_UPDATE_PROGRESS, &CChildView::OnCmUpdateProgress)
	ON_MESSAGE(CM_UPDATE_PROGRESS2, &CChildView::OnCmUpdateProgress2)
	ON_MESSAGE(CM_UPDATE_PROGRESS3, &CChildView::OnCmUpdateProgress3)
	ON_MESSAGE(CM_UPDATE_KEYPROGRESS1, &CChildView::OnCmUpdateKeyProgress1)
	ON_MESSAGE(CM_UPDATE_KEYPROGRESS2, &CChildView::OnCmUpdateKeyProgress2)
	//ON_COMMAND(ID_PROGRESS2, &CChildView::OnProgress2)
	ON_COMMAND(ID_BUTTON5, &CChildView::OnButton5)
	ON_UPDATE_COMMAND_UI(ID_BUTTON5, &CChildView::OnUpdateButton5)
	ON_COMMAND(ID_BUTTON3, &CChildView::OnButton3)
	ON_UPDATE_COMMAND_UI(ID_BUTTON3, &CChildView::OnUpdateButton3)
	ON_COMMAND(ID_CHECK3, &CChildView::OnCheck3)
	ON_UPDATE_COMMAND_UI(ID_CHECK3, &CChildView::OnUpdateCheck3)
	ON_UPDATE_COMMAND_UI(ID_PROGRESS3, &CChildView::OnUpdateProgress2)
	ON_COMMAND(ID_DIFFS_LIST, &CChildView::OnDiffslist)
	ON_UPDATE_COMMAND_UI(ID_DIFFS_LIST, &CChildView::OnUpdateDiffslist)
	ON_COMMAND(ID_SEL1, &CChildView::OnSel1)
	ON_COMMAND(ID_BUTTON6, &CChildView::OnButton6)
	ON_COMMAND(ID_PUT2FRONT, &CChildView::OnPut2front)
	ON_UPDATE_COMMAND_UI(ID_PUT2FRONT, &CChildView::OnUpdatePut2front)
	ON_WM_LBUTTONUP()
	ON_WM_RBUTTONUP()
	ON_UPDATE_COMMAND_UI(ID_COMBO2, &CChildView::OnUpdateCombo2)
	ON_COMMAND(ID_COMBO2, &CChildView::OnCombo2)
	ON_COMMAND(ID_SIMILARPAIRCHECKBOX, &CChildView::OnSimilarpaircheckbox)
	ON_UPDATE_COMMAND_UI(ID_SIMILARPAIRCHECKBOX, &CChildView::OnUpdateSimilarpaircheckbox)
	ON_COMMAND(ID_FINDREL_BTN, &CChildView::OnFindrelBtn)
	ON_COMMAND(ID_IDXCRT_BTN, &CChildView::OnIdxcrtBtn)
	ON_UPDATE_COMMAND_UI(ID_KEY_PROGRESS1, &CChildView::OnUpdateKeyProgress1)
	ON_UPDATE_COMMAND_UI(ID_KEY_PROGRESS2, &CChildView::OnUpdateKeyProgress2)
	ON_COMMAND(ID_USIDX_CHECK, &CChildView::OnUsidxCheck)
	ON_UPDATE_COMMAND_UI(ID_USIDX_CHECK, &CChildView::OnUpdateUsidxCheck)
	ON_UPDATE_COMMAND_UI(ID_ROWS1, &CChildView::OnUpdateRows1)
	ON_COMMAND(ID_ROWS1, &CChildView::OnRows1)
	ON_UPDATE_COMMAND_UI(ID_COLS1, &CChildView::OnUpdateCols1)
	ON_COMMAND(ID_COLS1, &CChildView::OnCols1)
	ON_UPDATE_COMMAND_UI(ID_ROWS2, &CChildView::OnUpdateRows2)
	ON_COMMAND(ID_ROWS2, &CChildView::OnRows2)
	ON_UPDATE_COMMAND_UI(ID_COLS2, &CChildView::OnUpdateCols2)
	ON_COMMAND(ID_COLS2, &CChildView::OnCols2)
END_MESSAGE_MAP()



// CChildView message handlers
/// <summary>
/// Pre-creates window.
/// </summary>
/// <param name="cs">The cs.</param>
/// <returns></returns>
BOOL CChildView::PreCreateWindow(CREATESTRUCT& cs)
{
	if (!CWnd::PreCreateWindow(cs))
		return FALSE;
	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;
	cs.style |= WS_VSCROLL | WS_HSCROLL;
	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS,
		::LoadCursor(NULL, IDC_ARROW), reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1), NULL);
	m_pRibbon = NULL;
	SYSTEM_INFO sysinfo;
	GetSystemInfo(&sysinfo);
	return TRUE;
}



/// <summary>
/// Called when [paint].
/// </summary>
void CChildView::OnPaint()
{
	CPaintDC dc(this); // device context for painting
#define TMPINT 64
	SendMessage(WM_ICONERASEBKGND, (WPARAM)dc.GetSafeHdc(), 0);
	CRect rect;
	GetClientRect(&rect);
	m_Clnt.w = (rect.Width());
	m_Clnt.h = (rect.Height());
	CPen pen1, pen2, pen3, pen4, pen5, pen6, pen7, pen8, pen9, pen10, pen11, pen12;
	CBrush brush0, brush1, brush2, brush3, brush4, brush5, brush6, brush7;
	pen1.CreatePen(PS_SOLID, 1, RGB(220, 220, 220));
	pen2.CreatePen(PS_SOLID, 1, RGB(200, 200, 200));
	pen3.CreatePen(PS_SOLID, 3, RGB(0, 0, 0));
	pen4.CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
	pen5.CreatePen(PS_SOLID, 1, RGB(100, 255, 100));
	pen6.CreatePen(PS_SOLID, 1, RGB(255, 255, 0));
	pen7.CreatePen(PS_SOLID, 2, RGB(0, 255, 0));
	pen8.CreatePen(PS_SOLID, 3, RGB(255, 0, 0));
	pen9.CreatePen(PS_SOLID, 2, RGB(155, 155, 255));
	pen10.CreatePen(PS_SOLID, 1, RGB(235, 235, 245));
	pen11.CreatePen(PS_SOLID, 1, RGB(255, 0, 0));
	pen12.CreatePen(PS_SOLID, 1, RGB(175, 0, 175));
	brush0.CreateSolidBrush(RGB(255, 255, 255));
	brush1.CreateSolidBrush(RGB(100, 255, 100));
	brush2.CreateSolidBrush(RGB(255, 255, 0));
	brush3.CreateSolidBrush(RGB(255, 127, 50));
	brush4.CreateSolidBrush(RGB(255, 80, 80));
	brush5.CreateSolidBrush(RGB(180, 180, 230));
	brush6.CreateSolidBrush(RGB(240, 240, 240));
	brush7.CreateSolidBrush(RGB(150, 200, 255));
	dc.SelectObject(&pen2);
	dc.SelectObject(&brush1);
	CFont font1, font2, font3, font4, font1B, font2B, font1C, font2C;
	font1.CreateFontW(16, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font2.CreateFontW(16, 0, 900, 900, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font3.CreateFontW(12, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font4.CreateFontW(30, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font1B.CreateFontW(16, 0, 0, 0, FW_EXTRABOLD, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font2B.CreateFontW(16, 0, 900, 900, FW_EXTRABOLD, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font1C.CreateFontW(12, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font2C.CreateFontW(12, 0, 900, 900, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
PaintCtx ctx;
ctx.pen1 = &pen1;   ctx.pen2 = &pen2;   ctx.pen3 = &pen3;   ctx.pen4 = &pen4;
ctx.pen5 = &pen5;   ctx.pen6 = &pen6;   ctx.pen7 = &pen7;   ctx.pen8 = &pen8;
ctx.pen9 = &pen9;   ctx.pen10 = &pen10; ctx.pen11 = &pen11; ctx.pen12 = &pen12;
ctx.brush0 = &brush0; ctx.brush1 = &brush1; ctx.brush2 = &brush2; ctx.brush3 = &brush3;
ctx.brush4 = &brush4; ctx.brush5 = &brush5; ctx.brush6 = &brush6; ctx.brush7 = &brush7;
ctx.font1 = &font1;   ctx.font2 = &font2;   ctx.font3 = &font3;   ctx.font4 = &font4;
ctx.font1B = &font1B; ctx.font2B = &font2B; ctx.font1C = &font1C; ctx.font2C = &font2C;
ctx.bnd_X_min = 1;
ctx.bnd_X_max = m_Table2.NumberOfColumns;
ctx.bnd_Y_min = 1;
ctx.bnd_Y_max = m_Table1.NumberOfColumns;
	paintInfoArea(dc, ctx);
	paintGridLines(dc, ctx);
	paintRowHeaders(dc, ctx);
	paintColumnHeaders(dc, ctx);
	paintMatrixCells(dc, ctx);
	paintSimilarityLines(dc, ctx);
	dc.SelectObject(&pen4);
	dc.MoveTo(0, OFFSET_Y + STEP_Y);
	dc.LineTo(m_Clnt.w, OFFSET_Y + STEP_Y);
	dc.MoveTo(OFFSET_X + STEP_X, 0);
	dc.LineTo(OFFSET_X + STEP_X, m_Clnt.h);
	m_bOnlyPcnt = false;
}



void CChildView::paintInfoArea(CDC& dc, PaintCtx& ctx)
{
	dc.SelectObject(ctx.font4);
	CString prcnt;
	prcnt = L"";
	if (!m_bToDisplaySimilarClms && M_CCell.x * M_CCell.y && m_nNatrixDone)
	{
		dc.SetBkMode(TRANSPARENT);
		if (M_CCell.x <= ctx.bnd_X_max && M_CCell.y <= ctx.bnd_Y_max)
		{
			long sameness = m_matrix.get(M_CCell.x, M_CCell.y);
			if (m_Table1.Columns[M_CCell.y] == m_Table2.Columns[M_CCell.x] && sameness < m_nEffMax)
				dc.SetTextColor(RGB(255, 0, 0));
			else
				dc.SetTextColor(RGB(0, 0, 0));
			prcnt.Format(L"\u0394:%i", m_nEffMax - sameness);
			dc.TextOutW(5, 20, prcnt);
			dc.SetTextColor(RGB(0, 255, 0));
			prcnt.Format(L"=:%i", sameness);
			dc.TextOutW(5, 50, prcnt);
			if (m_engine.isEmptyCol1(M_CCell.y) && m_engine.isEmptyCol2(M_CCell.x))
			{
				dc.SetTextColor(RGB(0, 0, 0));
				dc.TextOutW(5, 80, CMsg(IDS_EMPTY));
			}
		}
	}
	if (!m_bOnlyPcnt)
	{
		if (m_bNewFile1 && m_szFilename1)
		{
			dc.SetTextColor(RGB(120, 0, 130));
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(ctx.font1C);
			int index = ReverseFind(m_szFilename1, L"\\", -1) + 1;
			dc.TextOutW(2, 114, m_szFilename1.Mid(index, 22));
		}
		if (m_bNewFile2 && m_szFilename2)
		{
			dc.SetTextColor(RGB(120, 0, 130));
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(ctx.font2C);
			int index = ReverseFind(m_szFilename2, L"\\", -1) + 1;
			dc.TextOutW(112, 118, m_szFilename2.Mid(index, 22));
		}
	}
	if (M_CCell.x * M_CCell.y && m_bToDisplaySimilarClms)
	{
		if (M_CCell.x <= ctx.bnd_X_max && M_CCell.y <= ctx.bnd_Y_max)
		{
			dc.SetTextColor(RGB(50, 100, 250));
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(ctx.font1C);
			dc.TextOutW(5, 30, CMsg(IDS_KEY_SUITABILITY));
			dc.SelectObject(ctx.font4);
			prcnt.Format(L"~ %i%%", 100 * m_vecSimilaritiesAcrossTables[M_CCell.y].similarity / min(m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1, m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1));
			dc.TextOutW(15, 60, prcnt);
		}
	}
}



void CChildView::paintGridLines(CDC& dc, PaintCtx& ctx)
{
	dc.SelectObject(ctx.brush0);
	dc.SelectObject(ctx.pen2);
	dc.SelectObject(ctx.brush1);
	if (!m_bToDisplaySimilarClms)
	{
		for (int i_0 = OFFSET_X + STEP_X; i_0 < m_Clnt.w; i_0 += STEP_X)
		{
			dc.MoveTo(i_0, 0);
			dc.LineTo(i_0, m_Clnt.h);
		}
		for (int i_0 = OFFSET_Y + STEP_Y; i_0 < m_Clnt.h; i_0 += STEP_X)
		{
			dc.MoveTo(0, i_0);
			dc.LineTo(m_Clnt.w, i_0);
		}
	}
	dc.SelectObject(ctx.pen2);
	dc.SelectObject(ctx.brush1);
}



void CChildView::paintRowHeaders(CDC& dc, PaintCtx& ctx)
{
	int mx_y_adj;
	for (int mx_y = ctx.bnd_Y_min; mx_y <= ctx.bnd_Y_max; mx_y++)
	{
		bool cursor = false;
		mx_y_adj = mx_y + m_VisTopLeft.top;
		dc.SetBkMode(OPAQUE);
		if (isThisAKey(1, mx_y_adj))
		{
			if (mx_y_adj == m_OldCell.y)
			{
				dc.SelectObject(ctx.brush6);
			}
			else
			{
				if (mx_y_adj == M_CCell.y)
				{
					if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && (M_CCell.x > 0 || m_bToDisplaySimilarClms) && M_CCell.x <= m_Table2.NumberOfColumns)
					{
						cursor = true;
					}
					else
					{
						dc.SelectObject(ctx.brush0);
					}
				}
				else
				{
					dc.SelectObject(ctx.brush6);
				}
			}
		}
		else
		{
			dc.SelectObject(ctx.brush0);
		}
		if (mx_y_adj == M_CCell.y)
		{
			if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && (M_CCell.x > 0 || m_bToDisplaySimilarClms) && M_CCell.x <= m_Table2.NumberOfColumns)
			{
				 cursor = true;
			}
			else
			{
				if (isThisAKey(1, mx_y_adj))
				{
					dc.SelectObject(ctx.brush6);
				}
				else
				{
					dc.SelectObject(ctx.brush0);
				}
			}
		}
		dc.SelectObject(ctx.pen2);
		//else
			dc.Rectangle(0, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X, 1 + OFFSET_Y + mx_y * STEP_Y + STEP_Y);
		if (cursor)
		{
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(ctx.brush0);
			if (m_bToDisplaySimilarClms) dc.SelectObject(ctx.pen12); else dc.SelectObject(ctx.pen11);
			dc.Rectangle(2, 2 + OFFSET_Y + mx_y * STEP_Y, OFFSET_X + STEP_X - 1, -1 + OFFSET_Y + mx_y * STEP_Y + STEP_Y);
			dc.SetBkMode(OPAQUE);
			dc.SelectObject(ctx.brush0);
			dc.SelectObject(ctx.pen4);
		}
		if (m_nNatrixDone && !m_bOnlyPcnt && ((mx_y - m_VisTopLeft.top) > 0))
		{
			if (m_pbGreenClms1[mx_y])
			{
				dc.SelectObject(ctx.pen5);
				dc.SelectObject(ctx.brush1);
				dc.Ellipse(OFFSET_X, OFFSET_X + (mx_y - m_VisTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1, OFFSET_Y + STEP_Y + (mx_y - m_VisTopLeft.top) * STEP_Y);
			}
			if (m_engine.isEmptyCol1(mx_y))
			{
				dc.SelectObject(ctx.pen6);
				dc.SelectObject(ctx.brush2);
				dc.Ellipse(OFFSET_X, OFFSET_X + (mx_y - m_VisTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1, OFFSET_Y + STEP_Y + (mx_y - m_VisTopLeft.top) * STEP_Y);
			}
		}
		dc.SetBkMode(TRANSPARENT);
		if (isThisAKey(1, mx_y_adj))
		{
			dc.SelectObject(ctx.font1B);
			dc.SetTextColor(RGB(0, 0, 170));
		}
		else
		{
			dc.SelectObject(ctx.font1);
			dc.SetTextColor(RGB(0, 0, 0));
		}
		dc.TextOutW(2, OFFSET_Y + 5 + mx_y * STEP_Y, m_Table1.Columns[mx_y_adj]);
	}
}



void CChildView::paintColumnHeaders(CDC& dc, PaintCtx& ctx)
{
	int mx_x_adj;
	for (int mx_x = ctx.bnd_X_min; mx_x <= ctx.bnd_X_max; mx_x++)
	{
		bool cursor = false;
		mx_x_adj = mx_x + m_VisTopLeft.left;
		dc.SetBkMode(OPAQUE);
		if (isThisAKey(2, mx_x_adj))
		{
			if (mx_x_adj == m_OldCell.x)
			{
				dc.SelectObject(ctx.brush6);
			}
			else
			{
				if (mx_x_adj == M_CCell.x)
				{
					if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && (M_CCell.x > 0 || m_bToDisplaySimilarClms) && M_CCell.x <= m_Table2.NumberOfColumns)
					{
						cursor = true;
					}
					else
					{
						dc.SelectObject(ctx.brush0);
					}
				}
				else
				{
					dc.SelectObject(ctx.brush6);
				}
			}
		}
		else
		{
			dc.SelectObject(ctx.brush0);
		}
		if (mx_x_adj == M_CCell.x)
		{
			if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && (M_CCell.x > 0 || m_bToDisplaySimilarClms) && M_CCell.x <= m_Table2.NumberOfColumns)
			{
				 cursor = true;
			}
			else
			{
				if (isThisAKey(2, mx_x_adj))
				{
					dc.SelectObject(ctx.brush6);
				}
				else
				{
					dc.SelectObject(ctx.brush0);
				}
			}
		}
		dc.SelectObject(ctx.pen2);
		//else
			dc.Rectangle(OFFSET_X + mx_x * STEP_X, 0, 1 + OFFSET_X + mx_x * STEP_X + STEP_X, 1 + OFFSET_Y + STEP_Y);
		if (cursor)
		{
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(ctx.brush0);
			if (m_bToDisplaySimilarClms) dc.SelectObject(ctx.pen12); else dc.SelectObject(ctx.pen11);
			dc.Rectangle(2 + OFFSET_X + mx_x * STEP_X, 2, -1 + OFFSET_X + mx_x * STEP_X + STEP_X, OFFSET_Y + STEP_Y - 1);
			dc.SetBkMode(OPAQUE);
			dc.SelectObject(ctx.brush0);
			dc.SelectObject(ctx.pen4);
		}
		if (m_nNatrixDone && !m_bOnlyPcnt && ((mx_x - m_VisTopLeft.left) > 0))
		{
			if (m_pbGreenClms2[mx_x])
			{
				dc.SelectObject(ctx.pen5);
				dc.SelectObject(ctx.brush1);
				dc.Ellipse(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y, OFFSET_X + STEP_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
			}
			if (m_engine.isEmptyCol2(mx_x))
			{
				dc.SelectObject(ctx.pen6);
				dc.SelectObject(ctx.brush2);
				dc.Ellipse(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y, OFFSET_X + STEP_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
			}
		}
		dc.SetBkMode(TRANSPARENT);
		if (isThisAKey(2, mx_x_adj))
		{
			dc.SelectObject(ctx.font2B);
			dc.SetTextColor(RGB(0, 0, 170));
		}
		else
		{
			dc.SelectObject(ctx.font2);
			dc.SetTextColor(RGB(0, 0, 0));
		}
		dc.TextOutW(OFFSET_X + 5 + mx_x * STEP_X, -2 + OFFSET_Y + STEP_Y, m_Table2.Columns[mx_x_adj]);
	}
}



void CChildView::paintMatrixCells(CDC& dc, PaintCtx& ctx)
{
	dc.SelectObject(ctx.pen2);
	if (m_nNatrixDone && !m_bOnlyPcnt)
	{
		dc.SetBkMode(OPAQUE);
		dc.SelectObject(ctx.font3);
		int valSimil;
		CString strSimil;
		int mx_y_adj, mx_x_adj;
		if (m_nEffMax)
		{
			for (int mx_y = ctx.bnd_Y_min; mx_y <= ctx.bnd_Y_max - m_VisTopLeft.top; mx_y++)
			{
				for (int mx_x = ctx.bnd_X_min; mx_x <= ctx.bnd_X_max - m_VisTopLeft.left; mx_x++)
				{
					dc.SelectObject(ctx.pen2);
					mx_y_adj = mx_y + m_VisTopLeft.top;
					mx_x_adj = mx_x + m_VisTopLeft.left;
					valSimil = m_matrix.get(mx_x_adj, mx_y_adj) * 100 / m_nEffMax;
					strSimil.Format(L"%i", valSimil);
					strSimil += L"%";
					dc.SetBkMode(OPAQUE);
					if (!m_bSameNames || (m_Table1.Columns[mx_y_adj] == m_Table2.Columns[mx_x_adj]))
					{
						if (valSimil == 100)
						{
							if (m_engine.isEmptyCol1(mx_y_adj) || m_engine.isEmptyCol2(mx_x_adj))
							{
								dc.SelectObject(ctx.brush2);
							}
							else
							{
								dc.SelectObject(ctx.brush1);
							}
						}
						if (valSimil < 100)
						{
							if (valSimil > m_nSldr)
							{
								dc.SelectObject(ctx.brush4);
							}
							else
							{
								if (isThisAKey(1, mx_y_adj) || isThisAKey(2, mx_x_adj))
								{
									dc.SelectObject(ctx.brush6);
								}
								else
								{
									dc.SelectObject(ctx.brush0);
								}
							}
						}
					}
					else
					{
						if (isThisAKey(1, mx_y_adj) || isThisAKey(2, mx_x_adj))
						{
							dc.SelectObject(ctx.brush6);
						}
						else
						{
							dc.SelectObject(ctx.brush0);
						}
					}
					if (mx_y_adj == m_CClickedCell.y && mx_x_adj == m_CClickedCell.x)
					{
						dc.SelectObject(ctx.brush6);
					}
					dc.Rectangle(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X + mx_x * STEP_X, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y);
					dc.SetBkMode(TRANSPARENT);
					if (m_matrix.isMarked(mx_x_adj, mx_y_adj))
					{
						dc.SelectObject(ctx.pen4);
						dc.MoveTo(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y);
						dc.LineTo(OFFSET_X + STEP_X + mx_x * STEP_X, OFFSET_Y + STEP_Y + mx_y * STEP_Y);
						dc.MoveTo(OFFSET_X + STEP_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y);
						dc.LineTo(OFFSET_X + mx_x * STEP_X, OFFSET_Y + STEP_Y + mx_y * STEP_Y);
						dc.SelectObject(ctx.pen2);
					}
					if (m_bToDisplaySimilarClms && m_vecSimilaritiesAcrossTables[mx_y_adj].clm2 == mx_x_adj)
					{
						dc.SetBkMode(TRANSPARENT);
						dc.SelectObject(&m_KeyCurvePen);
						dc.SelectObject(ctx.brush7);
						dc.Rectangle(OFFSET_X + (mx_x)* STEP_X + 1, OFFSET_Y + (mx_y)* STEP_Y + 1, OFFSET_X + STEP_X + (mx_x)* STEP_X, OFFSET_Y + STEP_Y + (mx_y)* STEP_Y);
					}
					dc.SetTextColor(RGB(0, 0, 0));
					dc.TextOutW(OFFSET_X + mx_x * STEP_X + 1, OFFSET_Y + mx_y * STEP_Y + 7, strSimil);
				}
			}
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(GetStockObject(NULL_BRUSH));
			dc.SelectObject(ctx.pen3);
			for (int mx_y = ctx.bnd_Y_min; mx_y <= ctx.bnd_Y_max - m_VisTopLeft.top; mx_y++)
			{
				for (int mx_x = ctx.bnd_X_min; mx_x <= ctx.bnd_X_max - m_VisTopLeft.left; mx_x++)
				{
					mx_y_adj = mx_y + m_VisTopLeft.top;
					mx_x_adj = mx_x + m_VisTopLeft.left;
					if (m_Table1.Columns[mx_y_adj] == m_Table2.Columns[mx_x_adj])
					{
						dc.Rectangle(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X + mx_x * STEP_X, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y);
					}
					if (mx_y_adj == m_nOldy && mx_x_adj == m_nOldx)
					{
						dc.SelectObject(ctx.pen9);
						dc.Rectangle(OFFSET_X + mx_x * STEP_X + 3, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y - 4, 1 + OFFSET_X + STEP_X + mx_x * STEP_X - 2, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y - 2);
						dc.SelectObject(ctx.pen3);
					}
				}
			}
			dc.SelectObject(ctx.pen2);
		}
		m_bOnlyPcnt = false;
	}
}



void CChildView::paintSimilarityLines(CDC& dc, PaintCtx& ctx)
{
	if (m_bToDisplaySimilarClms)
	{
		int mx_x, mx_y = 0;
		long maxHit = m_vecSimilaritiesAcrossTablesSorted[1].similarity;
		for (int s_i = m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder; s_i >= 0; s_i--)
		{
			mx_y = m_vecSimilaritiesAcrossTablesSorted[s_i].clm1;
			mx_x = m_vecSimilaritiesAcrossTablesSorted[s_i].clm2;
			if ((mx_y * mx_x > 0) && ((mx_y - m_VisTopLeft.top) * (mx_x - m_VisTopLeft.left) > 0))
			{
				dc.SelectObject(&m_SimsPens[255 * m_vecSimilaritiesAcrossTablesSorted[s_i].similarity / maxHit]);
				CPoint pt[4] = {
					CPoint(OFFSET_X + STEP_X + 1, OFFSET_Y + (mx_y - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
					CPoint(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X , OFFSET_Y + (mx_y - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
					CPoint(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X + STEP_X / 2, OFFSET_Y + (mx_y - m_VisTopLeft.top) * STEP_Y),
					CPoint(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X + STEP_X / 2 , OFFSET_Y + STEP_Y)
				};
				dc.PolyBezier(pt, 4);
			}
		}
		for (int s_i = m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder; s_i >= 0; s_i--)
		{
			mx_y = m_vecSimilaritiesAcrossTablesSorted[s_i].clm1;
			mx_x = m_vecSimilaritiesAcrossTablesSorted[s_i].clm2;
			if ((mx_y * mx_x  > 0) && ((mx_y - m_VisTopLeft.top) * (mx_x - m_VisTopLeft.left) > 0))
			{
				if (isThisAKey(1, mx_y) && isThisAKey(2, mx_x))
				{
					dc.SelectObject(&m_SimsPens[255 * m_vecSimilaritiesAcrossTablesSorted[s_i].similarity / maxHit]);
					CPoint pt[4] = {
						CPoint(OFFSET_X + STEP_X + 1, OFFSET_Y + (mx_y - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
						CPoint(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X , OFFSET_Y + (mx_y - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
						CPoint(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X + STEP_X / 2, OFFSET_Y + (mx_y - m_VisTopLeft.top) * STEP_Y),
						CPoint(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X + STEP_X / 2 , OFFSET_Y + STEP_Y)
					};
					dc.SelectObject(&m_KeyCurvePen);
					dc.PolyBezier(pt, 4);
				}
			}
		}
	}
}



/// <summary>
/// Called when [pick first file].
/// </summary>
void CChildView::OnPickFirstFile()
{
	m_bNewFile1 = false;
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_TILL_IN_EXCEL)); // CMsg(IDS_WAIT_TILL_IN_EXCEL)
	CString fileName;
	wchar_t* p = fileName.GetBuffer(FILE_LIST_BUFFER_SIZE);
	CFileDialog dlgFile(TRUE);
	OPENFILENAME& ofn = dlgFile.GetOFN();
	//ofn.Flags |= OFN_ALLOWMULTISELECT; // for future scalability
	ofn.lpstrFile = p;
	ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
	dlgFile.DoModal();
	fileName.ReleaseBuffer();
	wchar_t* pBufEnd = p + FILE_LIST_BUFFER_SIZE - 2;
	wchar_t* start = p;
	while ((p < pBufEnd) && (*p))
		p++;
	if (p > start)
	{
		_tprintf(CMsg(IDS_PATH_TO_FILE), start); // CMsg(IDS_PATH_TO_FILE)
		p++;
		int fileCount = 1;
		while ((p < pBufEnd) && (*p))
		{
			start = p;
			while ((p < pBufEnd) && (*p))
				p++;
			if (p > start)
				_tprintf(_T("%2d. %s\r\n"), fileCount, start);
			p++;
			fileCount++;
		}
	}
	m_excel1.closeBook();
	if (!(CString(fileName) == L""))
	{
		for (int i = 0; i < 255; i++)
		{
			m_Table1.Columns[i] = "";
		}
		if (m_excel1.openFile(fileName, m_App))
		{
			m_pSheetCombo1->RemoveAllItems();
			CWorksheets& sheets = m_excel1.getSheets();
			for (int i = 1; i <= sheets.get_Count(); i++)
			{
				if (CWorksheet tempSheet = sheets.get_Item(COleVariant((short)i)))
				{
					m_pSheetCombo1->AddItem(tempSheet.get_Name());
				}
				else
				{
					break;
				}
			}
		}
	}
	m_szFilename1 = fileName;
	m_bNewFile1 = true;
	m_nUiToBeRefreshed = 3;
	if (m_nNatrixDone > 0)
	{
		m_matrix.clear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
		m_nNatrixDone = 0;
		m_OldCell.x = 0;
		m_OldCell.y = 0;
	}
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED)); // CMsg(IDS_FILE_SUCCESFULLY_LOADED)
	deleteAllKeys();
	this->Invalidate();
}



/// <summary>
/// Called when [pick second file].
/// </summary>
void CChildView::OnPickSecondFile()
{
	m_bNewFile2 = false;
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_TILL_IN_EXCEL)); // // CMsg(IDS_WAIT_TILL_IN_EXCEL)
	CString fileName;
	wchar_t* p = fileName.GetBuffer(FILE_LIST_BUFFER_SIZE);
	CFileDialog dlgFile(TRUE);
	OPENFILENAME& ofn = dlgFile.GetOFN();
	//ofn.Flags |= OFN_ALLOWMULTISELECT; // for future scalability
	ofn.lpstrFile = p;
	ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
	dlgFile.DoModal();
	fileName.ReleaseBuffer();
	wchar_t* pBufEnd = p + FILE_LIST_BUFFER_SIZE - 2;
	wchar_t* start = p;
	while ((p < pBufEnd) && (*p))
		p++;
	if (p > start)
	{
		_tprintf(CMsg(IDS_PATH_TO_FILE), start); // CMsg(IDS_PATH_TO_FILE)
		p++;
		int fileCount = 1;
		while ((p < pBufEnd) && (*p))
		{
			start = p;
			while ((p < pBufEnd) && (*p))
				p++;
			if (p > start)
				_tprintf(_T("%2d. %s\r\n"), fileCount, start);
			p++;
			fileCount++;
		}
	}
	m_excel2.closeBook();
	if (!(CString(fileName) == L""))
	{
		for (int i = 0; i < 255; i++)
		{
			m_Table2.Columns[i] = "";
		}
		if (m_excel2.openFile(fileName, m_App))
		{
			m_pSheetCombo2->RemoveAllItems();
			CWorksheets& sheets = m_excel2.getSheets();
			for (int i = 1; i <= sheets.get_Count(); i++)
			{
				if (CWorksheet tempSheet = sheets.get_Item(COleVariant((short)i)))
				{
					m_pSheetCombo2->AddItem(tempSheet.get_Name());
				}
				else
				{
					break;
				}
			}
		}
	}
	m_szFilename2 = fileName;
	m_bNewFile2 = true;
	m_nUiToBeRefreshed = 3;
	if (m_nNatrixDone > 0)
	{
		m_matrix.clear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
		m_nNatrixDone = 0;
	}
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED)); // CMsg(IDS_FILE_SUCCESFULLY_LOADED)
	deleteAllKeys();
	this->Invalidate();
}



/// <summary>
/// Called when [create matrix].
/// </summary>
void CChildView::OnCreateMatrix()
{
	if (m_bLockPrg1 || m_bLockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}
	m_nNatrixDone = 0;
	m_nPrereqDone = 0;
	if (areThereAnyKeys() == false)
	{
		MessageBox(CMsg(IDS_ATLEAST_ONE_KEY)); // CMsg(IDS_ATLEAST_ONE_KEY)
		return;
	}
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_COMPARISON_IN_PROGRESS)); // CMsg(IDS_COMPARISON_IN_PROGRESS)
	m_bWaitingForKeys = true;
	m_bKeys1done = false;
	m_bKeys2done = false;
	AfxBeginThread(CreateKeys1ThreadProc, this);
	AfxBeginThread(CreateKeys2ThreadProc, this);
	this->Invalidate();
	m_App.put_Visible(true);
	m_App.put_UserControl(TRUE);
}



/// <summary>
/// Called when [update pick first sheet].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdatePickFirstSheet(CCmdUI* pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSheetCombo1 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_PICK_FIRST_SHEET));
}



/// <summary>
/// Called when [update create matrix].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCreateMatrix(CCmdUI* pCmdUI)
{
	if (m_nUiToBeRefreshed)
	{
		pCmdUI->Enable(true);
		if (!(m_szFilename1 == ""))
		{
			CMFCRibbonBar* pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
			pRibbon->ForceRecalcLayout();
			this->GetTopLevelFrame()->Invalidate();
		}
		if (!(m_szFilename2 == ""))
		{
			CMFCRibbonBar* pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
			pRibbon->ForceRecalcLayout();
			this->GetTopLevelFrame()->Invalidate();
		}
		if (m_nUiToBeRefreshed > 0) m_nUiToBeRefreshed -= 1;
	}
}



/// <summary>
/// Called when [update filename1].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateFilename1(CCmdUI* pCmdUI)
{
	if (m_nUiToBeRefreshed)
	{
		if (!(m_szFilename1 == ""))
		{
			CString s = m_szFilename1;
			int origLen = s.GetLength();
			s = s.Right(20);
			s = (CString)CMsg(IDS_1ST_FILE) + (origLen > 20 ? ".." : "") + s; // CMsg(IDS_1ST_FILE)
			pCmdUI->SetText(s);
			pCmdUI->Enable(true);
			this->GetTopLevelFrame()->Invalidate();
		}
		else
		{
			pCmdUI->Enable(false);
		}
		if (m_nUiToBeRefreshed > 0) m_nUiToBeRefreshed -= 1;
	}
}



/// <summary>
/// Called when [update filename2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateFilename2(CCmdUI* pCmdUI)
{
	if (m_nUiToBeRefreshed)
	{
		if (!(m_szFilename2 == ""))
		{
			CString s = m_szFilename2;
			int origLen = s.GetLength();
			s = s.Right(20);
			s = (CString)CMsg(IDS_2ND_FILE) + (origLen > 20 ? ".." : "") + s; // CMsg(IDS_2ND_FILE)
			pCmdUI->SetText(s);
			pCmdUI->Enable(true);
			this->GetTopLevelFrame()->Invalidate();
		}
		else
		{
			pCmdUI->Enable(false);
		}
		if (m_nUiToBeRefreshed > 0) m_nUiToBeRefreshed -= 1;
	}
}



/// <summary>
/// Called when [mouse wheel].
/// </summary>
/// <param name="nFlags">The n flags.</param>
/// <param name="zDelta">The z delta.</param>
/// <param name="pt">The pt.</param>
/// <returns></returns>
BOOL CChildView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	int nDelta;
	nDelta = (zDelta / STEP_Y) * STEP_Y / (-5);
	int nScrollPos = m_nVScrollPos + nDelta;
	int nMaxPos = m_nViewHeight - m_nVPageSize;
	if (nScrollPos < 0)
		nDelta = -m_nVScrollPos;
	else if (nScrollPos > nMaxPos)
		nDelta = nMaxPos - m_nVScrollPos;
	if (nDelta != 0) {
		m_nVScrollPos += nDelta;
		m_VisTopLeft.top = m_nVScrollPos / STEP_Y;
		SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);
		//rect.left = OFFSET_X + STEP_X;
		rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(0, -nDelta, &rect);
		this->Invalidate();
	}
	m_bOnlyPcnt = false;
	m_bForceNotOnlyPcnt = true;
	m_OldCell.x = M_CCell.x;
	m_OldCell.y = M_CCell.y;
	M_CCell.x = 0;
	M_CCell.y = 0;
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(L"");
	return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}



/// <summary>
/// Called when [update pick second sheet].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdatePickSecondSheet(CCmdUI* pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSheetCombo2 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_PICK_SECOND_SHEET));
}



/// <summary>
/// Called when [update progress1].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateProgress1(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_PROGRESS2));
}



/// <summary>
/// Called when [v scroll].
/// </summary>
/// <param name="nSBCode">The n sb code.</param>
/// <param name="nPos">The n position.</param>
/// <param name="pScrollBar">The p scroll bar.</param>
void CChildView::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	int nDelta;
	switch (nSBCode) {
	case SB_LINEUP:
		nDelta = -LINESIZE;
		break;
	case SB_PAGEUP:
		nDelta = -m_nVPageSize;
		break;
	case SB_THUMBTRACK:
		nDelta = (int)nPos - m_nVScrollPos;
		break;
	case SB_PAGEDOWN:
		nDelta = m_nVPageSize;
		break;
	case SB_LINEDOWN:
		nDelta = LINESIZE;
		break;
	default: // Ignore other scroll bar messages
		return;
	}
	nDelta = (nDelta / STEP_Y) * STEP_Y;
	//
	// Adjust the delta if adding it to the current scroll position would
	// cause an underrun or overrun.
	//
	int nScrollPos = m_nVScrollPos + nDelta;
	int nMaxPos = m_nViewHeight - m_nVPageSize;
	if (nScrollPos < 0)
		nDelta = -m_nVScrollPos;
	else if (nScrollPos > nMaxPos)
		nDelta = nMaxPos - m_nVScrollPos;
	//
	// Update the scroll position and scroll the window.
	//
	if (nDelta != 0) {
		m_nVScrollPos += nDelta;
		m_VisTopLeft.top = m_nVScrollPos / STEP_Y;
		SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);
		//rect.left = OFFSET_X + STEP_X;
		rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(0, -nDelta, &rect);
		m_bOnlyPcnt = false;
		m_bForceNotOnlyPcnt = true;
		this->Invalidate();
	}
}



/// <summary>
/// Called when [h scroll].
/// </summary>
/// <param name="nSBCode">The n sb code.</param>
/// <param name="nPos">The n position.</param>
/// <param name="pScrollBar">The p scroll bar.</param>
void CChildView::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	int nDelta;
	switch (nSBCode) {
	case SB_LINELEFT:
		nDelta = -LINESIZE;
		break;
	case SB_PAGELEFT:
		nDelta = -m_nHPageSize;
		break;
	case SB_THUMBTRACK:
		nDelta = (int)nPos - m_nHScrollPos;
		break;
	case SB_PAGERIGHT:
		nDelta = m_nHPageSize;
		break;
	case SB_LINERIGHT:
		nDelta = LINESIZE;
		break;
	default: // Ignore other scroll bar messages
		return;
	}
	nDelta = (nDelta / STEP_X) * STEP_X;
	//
	// Adjust the delta if adding it to the current scroll position would
	// cause an underrun or overrun.
	//
	int nScrollPos = m_nHScrollPos + nDelta;
	int nMaxPos = m_nViewWidth - m_nHPageSize;
	if (nScrollPos < 0)
		nDelta = -m_nHScrollPos;
	else if (nScrollPos > nMaxPos)
		nDelta = nMaxPos - m_nHScrollPos;
	//
	// Update the scroll position and scroll the window.
	//
	if (nDelta != 0) {
		m_nHScrollPos += nDelta;
		m_VisTopLeft.left = m_nHScrollPos / STEP_X;
		SetScrollPos(SB_HORZ, m_nHScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);
		rect.left = OFFSET_X + STEP_X;
		//rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(-nDelta, 0, &rect);
		this->Invalidate();
	}
}







/// <summary>
/// Called when [pick first sheet].
/// </summary>
void CChildView::OnPickFirstSheet()
{
	int tmpWSN = m_pSheetCombo1->GetCurSel() + 1;
	CString tmpWSS = m_pSheetCombo1->GetEditText();
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_PRELIM_CHK)); // CMsg(IDS_WAIT_PRELIM_CHK)
	if (tmpWSN > 0)
	{
		m_Table1.WorkSheetNumber = tmpWSN;
		long iRows;
		long iCols;
		m_excel1.selectSheet(tmpWSS, iRows, iCols);
		m_Table1.MaxNumberOfRows = iRows;
		m_Table1.MaxNumberOfCols = iCols;
		m_Table1.NumberOfColumns = iCols;
		m_Table1.NumberOfRows = iRows;
		m_Table1.RowWithNames = 1;
		CString tmps;
		tmps.Format(_T("%d"), 1);
		m_pSpinner1_Names->SetEditText(tmps);
		m_Table1.RowWithNames = 1;
		tmps.Format(_T("%d"), 2);
		m_pSpinner1_Fdata->SetEditText(tmps);
		m_Table1.FirstRowWithData = 2;
		tmps.Format(_T("%d"), m_Table1.NumberOfRows);
		m_pRows1->SetEditText(tmps);
		tmps.Format(_T("%d"), m_Table1.NumberOfColumns);
		m_pCols1->SetEditText(tmps);
		updateCombos1();
		m_nCellWidth = STEP_X;
		m_nCellHeight = STEP_Y;
		m_nRibbonWidth = 0;
		m_nViewWidth = STEP_X + OFFSET_X + ((m_Table2.NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
		m_nViewHeight = STEP_Y + OFFSET_Y + m_nCellHeight * (m_Table1.NumberOfColumns + 1);
		SCROLLINFO si;
		si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
		si.nMin = 0;
		si.nMax = m_nViewHeight - 1;
		si.nPos = m_nVScrollPos;
		si.nPage = m_nVPageSize;
		SetScrollInfo(SB_VERT, &si, TRUE);
		this->Invalidate();
		m_nNatrixDone = false;
		deleteAllKeys();
		if (m_nNatrixDone > 0)
		{
			m_matrix.clear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
			m_nNatrixDone = 0;
			m_OldCell.x = 0;
			m_OldCell.y = 0;
			M_CCell.x = 0;
			M_CCell.y = 0;
		}
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED)); // CMsg(IDS_DATA_VERIFIED)
		m_engine.setTables(m_Table1, m_Table2);
		AfxBeginThread(makePrereq1ThreadProc, this);
	}
}



/// <summary>
/// Called when [spin1 names].
/// </summary>
void CChildView::OnSpin1Names()
{
	CString tmps = m_pSpinner1_Names->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 1) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	m_pSpinner1_Names->SetEditText(tmps);
	m_Table1.RowWithNames = tmpi;
	updateCombos1();
	this->Invalidate();
}



/// <summary>
/// Called when [update spin1 names].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateSpin1Names(CCmdUI* pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner1_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN1_NAMES));
}



/// <summary>
/// Called when [update spin1 fdata].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateSpin1Fdata(CCmdUI* pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner1_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN1_FDATA));
}



/// <summary>
/// Called when [spin1 fdata].
/// </summary>
void CChildView::OnSpin1Fdata()
{
	CString tmps = m_pSpinner1_Fdata->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 2) tmpi = 1; 
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	m_pSpinner1_Fdata->SetEditText(tmps);
	m_Table1.FirstRowWithData = tmpi;
	m_engine.invalidatePrereq1();
}



/// <summary>
/// Updates the combos1.
/// </summary>
void CChildView::updateCombos1()
{
	long index[2];
	CString szdata;
	COleVariant vData;
	for (int i = 1; i <= m_Table1.NumberOfColumns; i++)
	{
		// Loop through the data and report the contents.
		szdata = m_excel1.getCellValue(i, m_Table1.RowWithNames);
		if (szdata == "") szdata = CMsg(IDS_NO_NAME);
		for (int i1 = 1; i1 < i; i1++)
		{
			if (szdata == m_Table1.Columns[i1])
			{
				CString s;
				s.Format(L"[%i]", i);
				szdata += s;
				break;
			}
		}
		m_Table1.Columns[i] = szdata;
	}
}



/// <summary>
/// Called when [pick second sheet].
/// </summary>
void CChildView::OnPickSecondSheet()
{
	int tmpWSN = m_pSheetCombo2->GetCurSel() + 1;
	CString tmpWSS = m_pSheetCombo2->GetEditText();
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)); // CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)
	if (tmpWSN > 0)
	{
		m_Table2.WorkSheetNumber = tmpWSN;
		long iRows;
		long iCols;
		m_excel2.selectSheet(tmpWSS, iRows, iCols);
		m_Table2.MaxNumberOfRows = iRows;
		m_Table2.MaxNumberOfCols = iCols;
		m_Table2.NumberOfColumns = iCols;
		m_Table2.NumberOfRows = iRows;
		m_Table2.RowWithNames = 1;
		CString tmps;
		tmps.Format(_T("%d"), 1);
		m_pSpinner2_Names->SetEditText(tmps);
		m_Table2.RowWithNames = 1;
		tmps.Format(_T("%d"), 2);
		m_pSpinner2_Fdata->SetEditText(tmps);
		m_Table2.FirstRowWithData = 2;
		tmps.Format(_T("%d"), m_Table2.NumberOfRows);
		m_pRows2->SetEditText(tmps);
		tmps.Format(_T("%d"), m_Table2.NumberOfColumns);
		m_pCols2->SetEditText(tmps);
		m_nCellWidth = STEP_X;
		m_nCellHeight = STEP_Y;
		m_nRibbonWidth = 0;
		m_nViewWidth = STEP_X + OFFSET_X + ((m_Table2.NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
		m_nViewHeight = STEP_Y + OFFSET_Y + m_nCellHeight * (m_Table1.NumberOfColumns + 1);
		SCROLLINFO si;
		si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
		si.nMin = 0;
		si.nMax = m_nViewWidth - 1;
		si.nPos = m_nHScrollPos;
		si.nPage = m_nHPageSize;
		SetScrollInfo(SB_HORZ, &si, TRUE);
		deleteAllKeys();
		if (m_nNatrixDone > 0)
		{
			m_matrix.clear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
			m_nNatrixDone = 0;
			m_OldCell.x = 0;
			m_OldCell.y = 0;
			M_CCell.x = 0;
			M_CCell.y = 0;
		}
		updateCombos2();
		this->Invalidate();
		m_nNatrixDone = false;
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED)); // CMsg(IDS_DATA_VERIFIED)
		m_engine.setTables(m_Table1, m_Table2);
		AfxBeginThread(makePrereq2ThreadProc, this);
	}
}



/// <summary>
/// Updates the combos2.
/// </summary>
void CChildView::updateCombos2()
{
	long index[2];
	CString szdata;
	COleVariant vData;
	for (int i = 1; i <= m_Table2.NumberOfColumns; i++)
	{
		szdata = m_excel2.getCellValue(i, m_Table2.RowWithNames);
		if (szdata == "") szdata = CMsg(IDS_NO_NAME);
		for (int i1 = 1; i1 < i; i1++)
		{
			if (szdata == m_Table2.Columns[i1])
			{
				CString s;
				s.Format(L"[%i]", i);
				szdata += s;
				break;
			}
		}
		m_Table2.Columns[i] = szdata;
	}
}



/// <summary>
/// Called when [update spin2 fdata].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateSpin2Fdata(CCmdUI* pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner2_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN2_FDATA));
}



/// <summary>
/// Called when [spin2 fdata].
/// </summary>
void CChildView::OnSpin2Fdata()
{
	CString tmps = m_pSpinner2_Fdata->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 2) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	m_pSpinner2_Fdata->SetEditText(tmps);
	m_Table2.FirstRowWithData = tmpi;
	m_engine.invalidatePrereq2();
}




/// <summary>
/// Called when [update spin2 names].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateSpin2Names(CCmdUI* pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner2_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN2_NAMES));
}



/// <summary>
/// Called when [spin2 names].
/// </summary>
void CChildView::OnSpin2Names()
{
	CString tmps = m_pSpinner2_Names->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 1) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	m_pSpinner2_Names->SetEditText(tmps);
	m_Table2.RowWithNames = tmpi;
	updateCombos2();
	this->Invalidate();
}



/// <summary>
/// Makes the character arr1.
/// </summary>




/// <summary>
/// Called when [l button double CLK].
/// </summary>
/// <param name="nFlags">The n flags.</param>
/// <param name="point">The point.</param>
void CChildView::OnLButtonDblClk(UINT nFlags, CPoint point)
{
	if (m_bLockPrg1 || m_bLockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}
	if (m_nNatrixDone && (M_CCell.y <= m_Table1.NumberOfColumns && M_CCell.x <= m_Table2.NumberOfColumns))
	{
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
		m_bLockPrg2 = true;
		int mx_X_max = m_Table2.NumberOfColumns;
		m_matrix.setMarked(M_CCell.x, M_CCell.y);
		this->Invalidate();
		AfxBeginThread(MyThreadProc3, this);
	}
	CWnd::OnLButtonDblClk(nFlags, point);
}




/// <summary>
/// Checks the emptiness1.
/// </summary>




/// <summary>
/// Checks the keys uniqueness1.
/// </summary>
/// <returns></returns>
bool CChildView::checkKeysUniqueness1()
{
	m_bLockPrg1 = true;
	bool result = m_engine.checkKeysUniqueness1();
	m_bLockPrg1 = false;
	return result;
}



/// <summary>
/// Checks the keys uniqueness2.
/// </summary>
/// <returns></returns>
bool CChildView::checkKeysUniqueness2()
{
	m_bLockPrg2 = true;
	bool result = m_engine.checkKeysUniqueness2();
	m_bLockPrg2 = false;
	return result;
}



/// <summary>
/// First pass (main algorithm).
/// </summary>
void CChildView::firstPass()
{
	m_bLockPrg1 = true;
	m_engine.setTables(m_Table1, m_Table2);
	m_engine.firstPass(m_matrix, m_bAutoMark, m_bIn2file,
	                   m_pbGreenClms1, m_pbGreenClms2,
	                   m_nEffMax, m_bDoAutoMark);
	m_nNatrixDone++;
	m_bLockPrg1 = false;
}



/// <summary>
/// Creates the key arrays1.
/// </summary>
/// <returns></returns>
int CChildView::createKeyArrays1()
{
	m_engine.setTables(m_Table1, m_Table2);
	m_engine.m_bUseIndexes = m_bUseIndexes;
	m_bLockPrg1 = true;
	int result = m_engine.createKeyArrays1();
	m_bLockPrg1 = false;
	return result;
}



/// <summary>
/// Creates the key arrays2.
/// </summary>
/// <returns></returns>
int CChildView::createKeyArrays2()
{
	m_engine.setTables(m_Table1, m_Table2);
	m_engine.m_bUseIndexes = m_bUseIndexes;
	m_bLockPrg2 = true;
	int result = m_engine.createKeyArrays2();
	m_bLockPrg2 = false;
	return result;
}



/// <summary>
/// Called when [mouse move].
/// </summary>
/// <param name="nFlags">The n flags.</param>
/// <param name="point">The point.</param>
void CChildView::OnMouseMove(UINT nFlags, CPoint point)
{
	m_OldCell.x = M_CCell.x;
	m_OldCell.y = M_CCell.y;
	if (point.x > OFFSET_X + STEP_X)
	{
		M_CCell.x = (point.x - OFFSET_X) / STEP_X + m_VisTopLeft.left;
	}
	else
	{
		M_CCell.x = 0;
	}
	if (point.y > OFFSET_Y + STEP_Y)
	{
		M_CCell.y = (point.y - OFFSET_Y) / STEP_Y + m_VisTopLeft.top;
	}
	else
	{
		M_CCell.y = 0;
	}
	if (m_bToDisplaySimilarClms)
	{
		if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns)
		{
			int tmpCellx = m_vecSimilaritiesAcrossTables[M_CCell.y].clm2;
			if (tmpCellx > 0 && tmpCellx <= m_Table2.NumberOfColumns)
			{
				M_CCell.x = m_vecSimilaritiesAcrossTables[M_CCell.y].clm2;
			}
			else
			{
				M_CCell.x = M_CCell.y = 0;
			}
		}
		else
		{
			M_CCell.x = M_CCell.y = 0;
		}
	}
	if (M_CCell.x * M_CCell.y > 0)
	{
		CString s;
		CString sx, sy;
		sx.Format(L"%i", M_CCell.x);
		sy.Format(L"%i", M_CCell.y);
		sx = CMsg(IDS_COORDS);
		s.Format(CMsg(IDS_COORDS), M_CCell.y, M_CCell.x); // CMsg(IDS_COORDS)
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(s);
	}
	if (!(m_OldCell.x == M_CCell.x) || !(m_OldCell.y == M_CCell.y))
	{
		if (!m_bForceNotOnlyPcnt)
		{
			m_bOnlyPcnt = true;
		}
		else
		{
			m_bOnlyPcnt = false;
			m_bForceNotOnlyPcnt = false;
		}
		RECT rct;
		rct.left = 0; rct.top = 0; rct.right = OFFSET_X + STEP_X / 2; rct.bottom = OFFSET_Y + STEP_Y / 2;
		this->InvalidateRect(&rct, 1);
		if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && M_CCell.x > 0 && M_CCell.x <= m_Table2.NumberOfColumns)
		{
			rct.left = OFFSET_X + (M_CCell.x - m_VisTopLeft.left) * STEP_X + 1; rct.top = 2; rct.right = 1 + OFFSET_X + STEP_X + (M_CCell.x - m_VisTopLeft.left) * STEP_X; rct.bottom = OFFSET_Y + STEP_Y;
			this->InvalidateRect(&rct, 0);
			rct.left = 2; rct.top = OFFSET_Y + (M_CCell.y - m_VisTopLeft.top)  * STEP_Y + 1; rct.right = OFFSET_X + STEP_X; rct.bottom = 1 + OFFSET_Y + (M_CCell.y - m_VisTopLeft.top) * STEP_Y + STEP_Y;
			this->InvalidateRect(&rct, 0);
		}
		if (m_OldCell.y > 0 && m_OldCell.y <= m_Table1.NumberOfColumns && m_OldCell.x > 0 && m_OldCell.x <= m_Table2.NumberOfColumns)
		{
			rct.left = OFFSET_X + (m_OldCell.x  - m_VisTopLeft.left) * STEP_X + 1; rct.top = 2; rct.right = 1 + OFFSET_X + STEP_X + (m_OldCell.x  - m_VisTopLeft.left) * STEP_X; rct.bottom = OFFSET_Y + STEP_Y;
			this->InvalidateRect(&rct, 1);
			rct.left = 2; rct.top = OFFSET_Y + (m_OldCell.y  - m_VisTopLeft.top) * STEP_Y + 1; rct.right = OFFSET_X + STEP_X; rct.bottom = 1 + OFFSET_Y + (m_OldCell.y  - m_VisTopLeft.top) * STEP_Y + STEP_Y;
			this->InvalidateRect(&rct, 1);
		}
		//if (oldCell.x && oldCell.y)
		//{
		//	CString traces = L"";
		//	traces.Format(L"%i, %i, %i, %i\n", oldCell.x, oldCell.y, cCell.x, cCell.y);
		//	TRACE(traces);
		//}
		m_bOnlyPcnt = false;
		m_bForceNotOnlyPcnt = true;
		if (M_CCell.x * M_CCell.y == 0)
		{
			M_CCell.x = 0;
			M_CCell.y = 0;
		}
	}
	this->SetFocus();
}



/// <summary>
/// Called when [slider2].
/// </summary>
void CChildView::OnSlider2()
{
	m_nSldr = m_pSlider->GetPos();
	this->Invalidate();
	CString s;
	CString sx;
	s = m_szRsltTxt;
	sx.Format(CMsg(IDS_MARK_SUSP_INTERS), m_pSlider->GetPos()); // CMsg(IDS_MARK_SUSP_INTERS)
	s = sx + L" %";
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(s);
}



/// <summary>
/// Called when [update slider2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateSlider2(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSlider = DYNAMIC_DOWNCAST(CMFCRibbonSlider, m_pRibbon->FindByID(ID_SLIDER2));
	if (m_pSlider->GetPos() == 0)
		m_pSlider->SetPos(m_nSldr);
}



/// <summary>
/// Called when [check4].
/// </summary>
void CChildView::OnCheck4()
{
	m_bIn1file = !m_bIn1file;
}



/// <summary>
/// Called when [update check4].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCheck4(CCmdUI* pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pCmdUI->SetCheck(m_bIn1file);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pMarkIn1 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK4));
}



/// <summary>
/// Called when [check5].
/// </summary>
void CChildView::OnCheck5()
{
	m_bIn2file = !m_bIn2file;
}



/// <summary>
/// Called when [update check5].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCheck5(CCmdUI* pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pCmdUI->SetCheck(m_bIn2file);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pMarkIn2 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK5));
}



/// <summary>
/// Called when [button2].
/// </summary>
void CChildView::OnButton2()
{
	if (m_bLockPrg1 || m_bLockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}
/*	if (bestKeyComb.rating)
	{
		MessageBox(L"Vhodn� kombinace kl��� ji� byla nalezena"); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}  */
	if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns)
	{
		m_bWaitingForKeys = true;
		m_bKeysGathering1done = false;
		m_bKeysGathering2done = false;
		clearPossibleKeys();
		m_keyFinder.setTables(m_Table1, m_Table2);
		AfxBeginThread(SuggestKeys1ThreadProc, this);
		AfxBeginThread(SuggestKeys2ThreadProc, this);
	}
	else
	{
		MessageBox(CMsg(IDS_FRST_CHOOSE_DATA)); //CMsg(IDS_FRST_CHOOSE_DATA)
	}
}



/// <summary>
/// Converts the r1 c1.
/// </summary>
/// <param name="row">The row.</param>
/// <param name="clm">The CLM.</param>
/// <returns></returns>
CString CChildView::convertR1C1(int row, int clm)
{
	CString result;
	char chr1, chr2;
	chr2 = (clm + 25) % 26 + 65;
	result.Format(L"%i", row);
	result = chr2 + result;
	if (clm > 26)
	{
		chr1 = clm / 26 + 64;
		result = chr1 + result;
	}
	return result;
}



/// <summary>
/// Marks a difference in the XLS(X) file.
/// </summary>
/// <param name="row">The row.</param>
/// <param name="clm">The CLM.</param>
void CChildView::markIn1(int row, int clm)
{
	CString cnv = convertR1C1(row, clm);
	m_excel1.markCellRange(cnv, cnv, RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
}



/// <summary>
/// Marks a difference in the XLS(X) file.
/// </summary>
/// <param name="row">The row.</param>
/// <param name="clm">The CLM.</param>
void CChildView::markIn2(int row, int clm)
{
	CString cnv = convertR1C1(row, clm);
	m_excel2.markCellRange(cnv, cnv, RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
}



/// <summary>
/// Initializes the scroll bars.
/// </summary>
void CChildView::initScrollBars()
{
	SCROLLINFO ScrollInfo;
	ScrollInfo.cbSize = sizeof(ScrollInfo);     // size of this structure
	ScrollInfo.fMask = SIF_ALL;                 // parameters to set
	ScrollInfo.nMin = 0;                        // minimum scrolling position
	ScrollInfo.nMax = 100;                      // maximum scrolling position
	ScrollInfo.nPage = 20;                      // the page size of the scroll box
	ScrollInfo.nPos = 50;
	// initial position of the scroll box
	//ScrollInfo.nTrackPos = 0;                   // immediate position of a scroll box
	this->SetScrollInfo(SB_HORZ, &ScrollInfo);
}



/// <summary>
/// Called when [size].
/// </summary>
/// <param name="nType">Type of the n.</param>
/// <param name="cx">The cx.</param>
/// <param name="cy">The cy.</param>
void CChildView::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);
	//
	// Set the horizontal scrolling parameters.
	//
	int nHScrollMax = 0;
	m_nHScrollPos = m_nHPageSize = 0;
	if (cx < m_nViewWidth) {
		nHScrollMax = m_nViewWidth - 1;
		m_nHPageSize = cx;
		m_nHScrollPos = min(m_nHScrollPos, m_nViewWidth -
			m_nHPageSize - 1);
		m_VisTopLeft.left = 0;
	}
	SCROLLINFO si;
	si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
	si.nMin = 0;
	si.nMax = nHScrollMax;
	si.nPos = m_nHScrollPos;
	si.nPage = m_nHPageSize;
	SetScrollInfo(SB_HORZ, &si, TRUE);
	//
	// Set the vertical scrolling parameters.
	//
	int nVScrollMax = 0;
	m_nVScrollPos = m_nVPageSize = 0;
	if (cy < m_nViewHeight) {
		nVScrollMax = m_nViewHeight - 1;
		m_nVPageSize = cy;
		m_nVScrollPos = min(m_nVScrollPos, m_nViewHeight -
			m_nVPageSize - 1);
		m_VisTopLeft.top = 0;
	}
	si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
	si.nMin = 0;
	si.nMax = nVScrollMax;
	si.nPos = m_nVScrollPos;
	si.nPage = m_nVPageSize;
	SetScrollInfo(SB_VERT, &si, TRUE);
	m_bOnlyPcnt = false;
	//this->Invalidate(); // uncomment in case of problems with redrawing after RESIZE
	m_VisTopLeft.top = m_nVScrollPos / STEP_Y;
	SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
	RECT rect;
	GetClientRect(&rect);
	rect.top = OFFSET_Y + STEP_Y;
	ScrollWindow(0, 0, &rect);
	m_bOnlyPcnt = false;
	m_bForceNotOnlyPcnt = true;
	this->Invalidate();
}



/// <summary>
/// Called when [create].
/// </summary>
/// <param name="lpCreateStruct">The lp create structure.</param>
/// <returns></returns>
int CChildView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	CClientDC dc(this);
	m_nCellWidth = STEP_X;
	m_nCellHeight = STEP_Y;
	m_nRibbonWidth = 0;
	m_nViewWidth = STEP_X + OFFSET_X + ((m_Table2.NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
	m_nViewHeight = STEP_Y + OFFSET_Y + m_nCellHeight * (m_Table1.NumberOfColumns + 1);
	m_nSldr = 90;
	m_engine.init(GetSafeHwnd(), m_excel1, m_excel2, m_Table1, m_Table2);
	m_keyFinder.init(GetSafeHwnd(), m_excel1, m_excel2, m_Table1, m_Table2);
	return 0; 
}



/// <summary>
/// Called when [update progress2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateProgress2(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pProgressBar2 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_PROGRESS3));
	// Emergency update of the container for found differences
	m_pFoundDifferences = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_DIFFS_LIST));
	m_pToFront = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_PUT_TO_FRONT));
}



/// <summary>
/// Called when [update check2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCheck2(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_bVerifyKeys);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pVerifyKeys = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK2));
}



/// <summary>
/// Called when [check2].
/// </summary>
void CChildView::OnCheck2()
{
	m_bVerifyKeys = !m_bVerifyKeys;
}



/// <summary>
/// Called when [update button2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateButton2(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(true);
	//pCmdUI->SetText(m_bUseIndexes ? L"Sestavit kl��" : L"Naj�t kl��");
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pButton2 = DYNAMIC_DOWNCAST(CMFCRibbonButton, m_pRibbon->FindByID(ID_BUTTON2));
}



/// <summary>
/// Called when [check7].
/// </summary>
void CChildView::OnCheck7()
{
	m_bSameNames = !m_bSameNames;
	m_VisTopLeft.top = m_nVScrollPos / STEP_Y;
	SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
	RECT rect;
	GetClientRect(&rect);
	rect.top = OFFSET_Y + STEP_Y;
	ScrollWindow(0, 0, &rect);
	m_bOnlyPcnt = false;
	m_bForceNotOnlyPcnt = true;
	this->Invalidate();
}



/// <summary>
/// Called when [update check7].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCheck7(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_bSameNames);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSameNames = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK7));
}



/// <summary>
/// Mies the thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT MyThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->firstPass();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Called when [cm update progress].
/// </summary>
/// <param name="wParam">The w parameter.</param>
/// <param name="lParam">The l parameter.</param>
/// <returns></returns>
afx_msg LRESULT CChildView::OnCmUpdateProgress(WPARAM wParam, LPARAM lParam)
{
	if ((UINT)lParam > 99)
	{
		m_pProgressBar1->SetPos(0);
		m_bLockPrg1 = false;
		this->Invalidate();
		if (m_bDoAutoMark)
		{
			if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
			resolveAutoMark();
			if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
		}
		if ((UINT)lParam == 1000)
		{
			if (m_bWaitingForKeys)
			{
				m_bKeys1done = true;
				if (m_bKeys2done)
				{
					m_bWaitingForKeys = false;
					m_bKeys1done = false;
					m_bKeys2done = false;
					AfxBeginThread(MyThreadProc, this);
					if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
				}
			}
			else
			{
				if (m_nEffMax)
				{
					m_szRsltTxt.Format(CMsg(IDS_FOUND_KEYS_FROM_TOTAL), m_nEffMax, (m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1), (m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1)); // CMsg(IDS_FOUND_KEYS_FROM_TOTAL)
					if (g_pMainFrame) g_pMainFrame->updateStatusBar(m_szRsltTxt);
				}
			}
		}
		if ((UINT)lParam == 10000)
		{
			if (m_bWaitingForKeys)
			{
				m_bKeysGathering1done = true;
				if (m_bKeysGathering2done)
				{
					m_bWaitingForKeys = false;
					m_bKeysGathering1done = false;
					m_bKeysGathering2done = false;
					AfxBeginThread(MutualCheckThreadProc, this);
					BeginWaitCursor();
					if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING));
				}
			}
		}
		if ((UINT)lParam == 100000)
		{
			m_bWaitingForKeys = false;
			usePossibleKeys();	
			CString tmpS;
			tmpS.Format(CMsg(IDS_KEY_COMB_FOUND), m_keyFinder.getBestKeyComb().cnt); // CMsg(IDS_KEY_COMB_FOUND)
			MessageBox(tmpS);
			EndWaitCursor();
		}
		if ((UINT)lParam == 200000)
		{
			m_bWaitingForKeys = false;
			MessageBox(CMsg(IDS_INCOMPATBL_KEY_FOUND)); // CMsg(IDS_INCOMPATBL_KEY_FOUND)
			EndWaitCursor();
		}
	}
	else
	{
		m_pProgressBar1->SetPos((UINT)lParam);
	}
	return 0;
}



/// <summary>
/// Called when [cm update progress2].
/// </summary>
/// <param name="wParam">The w parameter.</param>
/// <param name="lParam">The l parameter.</param>
/// <returns></returns>
afx_msg LRESULT CChildView::OnCmUpdateProgress2(WPARAM wParam, LPARAM lParam)
{
	if ((UINT)lParam > 99)
	{
		m_pProgressBar2->SetPos(0);
		m_bLockPrg2 = false;
		if ((UINT)lParam == 1000)
		{
			this->Invalidate();
			if (m_bWaitingForKeys)
			{
				m_bKeys2done = true;
				if (m_bKeys1done)
				{
					m_bWaitingForKeys = false;
					m_bKeys1done = false;
					m_bKeys2done = false;
					AfxBeginThread(MyThreadProc, this);
					if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
				}
			}
		}
		if ((UINT)lParam == 20000)
		{
			if (m_bWaitingForKeys)
			{
				m_bKeysGathering2done = true;
				if (m_bKeysGathering1done)
				{
					m_bWaitingForKeys = false;
					m_bKeysGathering1done = false;
					m_bKeysGathering2done = false;
					AfxBeginThread(MutualCheckThreadProc, this);
					if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
				}
			}
		}
	}
	else
	{
		m_pProgressBar2->SetPos((UINT)lParam);
	}
	return 0;
}



/// <summary>
/// Called when [cm update progress3].
/// </summary>
/// <param name="wParam">The w parameter.</param>
/// <param name="lParam">The l parameter.</param>
/// <returns></returns>
afx_msg LRESULT CChildView::OnCmUpdateProgress3(WPARAM wParam, LPARAM lParam)
{
	HWND hWnd = this->GetSafeHwnd();
	long nor;
	int prgHlpr, prgHlpr0;
	prgHlpr = 0;
	prgHlpr0 = 0;
	CString fndDfrnc, fndDfrnc1, fndDfrnc2, selKey;
	int dfrnCntr = 0;
	long dfrncRow2;
	CString temps = L"";
	CString starts = L"";
	CString ends = L"";
	if ((UINT)lParam >100)
	{
		BeginWaitCursor();
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING));  // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		m_pProgressBar1->SetPos(0);
		m_pFoundDifferences->RemoveAllItems();
		m_pFoundDifferences->SetEditText(L"");
		{
			nor = m_Table1.NumberOfRows + 1;
			for (int i1 = 1; i1 < nor; i1++)
			{
				prgHlpr = (i1 * 100) / nor;
				if (prgHlpr > prgHlpr0)
				{
					SendMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
					prgHlpr0 = prgHlpr;
				}
				dfrncRow2 = m_pnFoundDifferences[i1];
				if (dfrncRow2 > 0)
				{
					if (++dfrnCntr < 500)
					{
						fndDfrnc1 = L"";
						fndDfrnc1.Format(L"(1r%i):", i1);
						fndDfrnc1 += m_excel1.getCellValue(m_nOldy, i1);
						fndDfrnc1 = fndDfrnc1.Left(26);
						fndDfrnc2 = L"";
						fndDfrnc2.Format(L"   (2r%i):", dfrncRow2);
						fndDfrnc2 += m_excel2.getCellValue(m_nOldx, dfrncRow2);
						fndDfrnc2 = fndDfrnc2.Left(26);
						selKey = L"";
						selKey.Format(L"%s%s   (key): %s", fndDfrnc1, fndDfrnc2, m_engine.getKeyStr1(i1));
						fndDfrnc = selKey.Left(54);
						//fndDfrnc = fndDfrnc1 + fndDfrnc2 + selKey;
						m_pFoundDifferences->AddItem((LPCTSTR)fndDfrnc);
					}
				}
				if (m_bIn1file)
				{
					if (m_pbMarkIn1Arr[i1])
					{
						if (starts == L"")
						{
							starts = convertR1C1(i1, m_nOldy);
						}
						ends = convertR1C1(i1, m_nOldy);
					}
					else
					{
						if (!(starts == L"") && !(ends == L""))
						{
							m_excel1.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
							starts = L"";
							ends = L"";
						}
					}
				}
			}
			if (m_bIn1file && !(starts == L"") && !(ends == L""))
			{
				m_excel1.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
				starts = L"";
				ends = L"";
			}
		}
		temps = L"";
		starts = L"";
		ends = L"";
		if (m_bIn2file)
		{
			nor = m_Table2.NumberOfRows + 1;
			for (int i2 = 1; i2 < nor; i2++)
			{
				prgHlpr = (i2 * 100) / nor;
				if (prgHlpr > prgHlpr0)
				{
					SendMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
					prgHlpr0 = prgHlpr;
				}
				if (m_pbMarkIn2Arr[i2])
				{
					if (starts == L"")
					{
						starts = convertR1C1(i2, m_nOldx);
					}
					ends = convertR1C1(i2, m_nOldx);
				}
				else
				{
					if (!(starts == L"") && !(ends == L""))
					{
						m_excel2.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
						starts = L"";
						ends = L"";
					}
				}
			}
			if (!(starts == L"") && !(ends == L""))
			{
				m_excel2.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
				starts = L"";
				ends = L"";
			}
		}
		SendMessage(CM_UPDATE_PROGRESS2, 0, 100);
		m_bLockPrg2 = false;
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE)); // CMsg(IDS_MARKING_DONE)
		EndWaitCursor();
		DrainMsgQueue();
	}
	else
	{
		m_pProgressBar1->SetPos((UINT)lParam);
	}
	return 0;
}



/// <summary>
/// Mies the thread proc2.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT MyThreadProc2(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->m_bUniqueKeys1 = false;
	pWnd->m_bUniqueKeys2 = false;
	int rslt;
	rslt = pWnd->createKeyArrays1();
	if (rslt == 1)
	{
		pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE)); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE)
		return 0;
	}
	pWnd->m_bUniqueKeys1 = true;
	rslt = pWnd->createKeyArrays2();
	if (rslt == 2)
	{
		pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE));// CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE)
		return 0;
	}
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Creates the keys1 thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT CreateKeys1ThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->m_bUniqueKeys1 = false;
	int rslt;
	rslt = pWnd->createKeyArrays1();
	if (rslt == 1)
	{
		CString s;
		NotUniqueKeys* nu = &pWnd->m_engine.m_NotUniqueKeys1;
		s.Format(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow, nu->secondRow); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS)
		pWnd->MessageBox(s);
		pWnd->m_bLockPrg1 = false;
		return 0;
	}
	pWnd->m_bUniqueKeys1 = true;
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Creates the keys2 thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT CreateKeys2ThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->m_bUniqueKeys2 = false;
	int rslt;
	rslt = pWnd->createKeyArrays2();
	if (rslt == 2)
	{
		CString s;
		NotUniqueKeys* nu = &pWnd->m_engine.m_NotUniqueKeys2;
		s.Format(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow, nu->secondRow); // CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS)
		pWnd->MessageBox(s);
		pWnd->m_bLockPrg2 = false;
		return 0;
	}
	pWnd->m_bUniqueKeys2 = true;
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Makes the prereq1 thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT makePrereq1ThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->m_engine.makePrereq1();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Makes the prereq2 thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT makePrereq2ThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->m_engine.makePrereq2();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Mies the thread proc3.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT MyThreadProc3(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->markInFiles();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Marks differences in XLS(X) files.
/// </summary>
void CChildView::markInFiles()
{
	m_bLockPrg2 = true;
	CString concatenatedKey1, concatenatedKey2;
	int prgHlpr = 0, prgHlpr0 = 0;
	int cx, cy;
	cx = M_CCell.x;
	cy = M_CCell.y;
	m_nOldy = cy;
	m_nOldx = cx;
	m_pbMarkIn1Arr.assign(m_Table1.NumberOfRows + 2, false);
	m_pbMarkIn2Arr.assign(m_Table2.NumberOfRows + 2, false);
	m_pnFoundDifferences.assign(m_Table1.NumberOfRows + 2, 0L);
	for (int i1 = 0; i1 <= m_Table1.NumberOfRows + 1; i1++)
	{
		m_pbMarkIn1Arr[i1] = false;
		m_pnFoundDifferences[i1] = 0;
	}
	for (int i2 = 0; i2 <= m_Table2.NumberOfRows + 1; i2++)
	{
		m_pbMarkIn2Arr[i2] = false;
	}
	long keyRow1, keyRow2;
	POSITION mapPos1;
	mapPos1 = m_engine.getMap1().GetStartPosition();
	int i1; // iterator for progress visualisation;
	i1 = m_Table1.FirstRowWithData - 1;
	while (mapPos1 != NULL)
	{
		i1++;
		prgHlpr0 = 100 * i1 / m_Table1.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS3, 0, prgHlpr);
		}
		m_engine.getMap1().GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);
		if (m_engine.getMap2().Lookup(concatenatedKey1, (long&)keyRow2))
		{
			if (!(m_excel1.getCellValue(cy, keyRow1) == m_excel2.getCellValue(cx, keyRow2)))
			{
				m_pnFoundDifferences[keyRow1] = keyRow2;
				if (m_bIn1file) m_pbMarkIn1Arr[keyRow1] = true; //markIn1(i1, cy);
				if (m_bIn2file) m_pbMarkIn2Arr[keyRow2] = true; //markIn2(i2, cx);
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS3, 0, 1000);
	m_bLockPrg2 = false;
}



/// <summary>
/// Called when [button5].
/// </summary>
void CChildView::OnButton5()
{
	COLORREF i = (int)m_pColorPicker1->GetSelectedItem();
	m_nChosenColor1 = i;
}



/// <summary>
/// Called when [update button5].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateButton5(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(true);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pColorPicker1 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_pRibbon->FindByID(ID_BUTTON5));
}



/// <summary>
/// Called when [button3].
/// </summary>
void CChildView::OnButton3()
{
	COLORREF i = (int)m_pColorPicker2->GetSelectedItem();
	m_nChosenColor2 = i;
}



/// <summary>
/// Called when [update button3].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateButton3(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(true);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pColorPicker2 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_pRibbon->FindByID(ID_BUTTON3));
}



/// <summary>
/// Called when [check3].
/// </summary>
void CChildView::OnCheck3()
{
	m_bAutoMark = !m_bAutoMark;
}



/// <summary>
/// Called when [update check3].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCheck3(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(true);
	pCmdUI->SetCheck(m_bAutoMark);
}



/// <summary>
/// Resolves the AutoMark.
/// </summary>
void CChildView::resolveAutoMark()
{
	m_bDoAutoMark = false;
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_DURING_MARKING_THREAD_BLOCKED));  // CMsg(IDS_DURING_MARKING_THREAD_BLOCKED)
	m_bLockPrg2 = true;
	HWND hWnd = this->GetSafeHwnd();
	int prgHlpr_x, prgHlpr0_x, prgHlpr_y, prgHlpr0_y;
	prgHlpr_x = 0;
	prgHlpr0_x = 0;
	prgHlpr_y = 0;
	prgHlpr0_y = 0;
	CString starts = L"";
	CString ends = L"";
	m_pProgressBar1->SetPos(0);
	BeginWaitCursor();
	for (int c1 = 1; c1 <= m_Table1.NumberOfColumns; c1++)
	{
		prgHlpr0_x = 90 * c1 / m_Table1.NumberOfColumns;
		if (prgHlpr0_x > prgHlpr_x)
		{
			prgHlpr_x = prgHlpr0_x;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr_x);
		}
		for (int c2 = 1; c2 <= m_Table2.NumberOfColumns; c2++)
		{
			if (m_Table1.Columns[c1] == m_Table2.Columns[c2])
			{
				if (m_bIn1file)
				{
					prgHlpr_y = 0;
					prgHlpr0_y = 0;
					for (long r1 = m_Table1.FirstRowWithData; r1 <= m_Table1.NumberOfRows; r1++)
					{
						prgHlpr0_y = 100 * r1 / m_Table1.NumberOfRows;
						if (prgHlpr0_y > prgHlpr_y + 10)
						{
							prgHlpr_y = prgHlpr0_y;
							PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr_y);
						}
						if (m_engine.getMainChar1(r1, c1) == 1)
						{
							if (starts == L"")
							{
								starts = convertR1C1(r1, c1);
							}
							ends = convertR1C1(r1, c1);
						}
						else
						{
							if (!(starts == L"") && !(ends == L""))
							{
								m_excel1.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
								starts = L"";
								ends = L"";
							}
						}
					}
					if (!(starts == L"") && !(ends == L""))
					{
						m_excel1.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
						starts = L"";
						ends = L"";
					}
				}
				starts = L"";
				ends = L"";
				if (m_bIn2file)
				{
					prgHlpr_y = 0;
					prgHlpr0_y = 0;
					for (long r2 = m_Table2.FirstRowWithData; r2 <= m_Table2.NumberOfRows; r2++)
					{
						prgHlpr0_y = 100 * r2 / m_Table2.NumberOfRows;
						if (prgHlpr0_y > prgHlpr_y + 10)
						{
							prgHlpr_y = prgHlpr0_y;
							PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr_y);
						}
						if (m_engine.getMainChar2(r2, c2) == 1)
						{
							if (starts == L"")
							{
								starts = convertR1C1(r2, c2);
							}
							ends = convertR1C1(r2, c2);
						}
						else
						{
							if (!(starts == L"") && !(ends == L""))
							{
								m_excel2.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
								starts = L"";
								ends = L"";
							}
						}
					}
					if (!(starts == L"") && !(ends == L""))
					{
						m_excel2.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
						starts = L"";
						ends = L"";
					}
				}
			}
		}
	}
	for (long r1 = m_Table1.FirstRowWithData; r1 <= m_Table1.NumberOfRows; r1++)
	{
		if (m_engine.isKeyMissing1(r1))
		{
			starts = convertR1C1(r1, 1);
			ends = convertR1C1(r1, m_Table1.NumberOfColumns);
			m_excel1.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
		}
	}
	for (long r2 = m_Table2.FirstRowWithData; r2 <= m_Table2.NumberOfRows; r2++)
	{
		if (m_engine.isKeyMissing2(r2)) // c1 - because we need it to run just once
		{
			starts = convertR1C1(r2, 1);
			ends = convertR1C1(r2, m_Table2.NumberOfColumns);
			m_excel2.markCellRange(starts, ends, RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
		}
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 100);
	PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
	m_bLockPrg2 = false;
	EndWaitCursor();
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE)); // CMsg(IDS_MARKING_DONE)
	DrainMsgQueue();
}



/// <summary>
/// Drain the MSG queue. // TODO explain
/// </summary>
void CChildView::DrainMsgQueue(void)
{
	MSG     msg = { 0 };
	HWND hWnd = this->GetSafeHwnd();
	while (PeekMessage(&msg, hWnd, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE));
}



/// <summary>
/// Called when [diffslist].
/// </summary>
void CChildView::OnDiffslist()
{
	// there is no required answer for this event - at least for now
}


/// <summary>
/// Called when [update diffslist].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateDiffslist(CCmdUI* pCmdUI)
{
	// there is no required answer for this event - at least for now
}



/// <summary>
/// Called when [sel1].
/// </summary>
void CChildView::OnSel1()
{
	long row;
	long column;
	row = rowFromCombo();
	if (row > 0)
	{
		column = m_nOldy;
		CString cnv = convertR1C1(row, column);
		m_excel1.selectAndActivateCell(cnv);
		if (m_bToFront)
		{
			m_App.put_Interactive(true);
			HWND ehWnd = (HWND)m_App.get_Hwnd();
			::PostMessage(ehWnd, WM_SHOWWINDOW, SW_RESTORE, 0);
			::SetForegroundWindow(ehWnd);
		}
	}
}



/// <summary>
/// gets row from combo.
/// </summary>
/// <returns></returns>
int CChildView::rowFromCombo()
{
	if (m_pFoundDifferences->GetCurSel() > -1)
	{
		CString s;
		s = m_pFoundDifferences->GetEditText();
		int bct1, bct2;
		bct1 = s.Find('(', 0) + 3;
		bct2 = s.Find(')', bct1 + 1);
		CString d;
		d = s.Mid(bct1, bct2 - bct1);
		int rslt;
		rslt = _tstoi(d);
		return rslt;
	}
	return 0;
}



/// <summary>
/// Called when [button6].
/// </summary>
void CChildView::OnButton6()
{
	long row;
	long column;
	row = rowFromCombo();
	if (row > 0)
	{
		row = m_pnFoundDifferences[row];
		column = m_nOldx;
		CString cnv = convertR1C1(row, column);
		m_excel2.selectAndActivateCell(cnv);
		if (m_bToFront)
		{
			m_App.put_Interactive(true);
			HWND ehWnd = (HWND)m_App.get_Hwnd();
			::PostMessage(ehWnd, WM_SHOWWINDOW, SW_RESTORE, 0);
			::SetForegroundWindow(ehWnd);
		}
	}
}



/// <summary>
/// Called when [put2front].
/// </summary>
void CChildView::OnPut2front()
{
	m_bToFront = !m_bToFront;
}



/// <summary>
/// Called when [update put2front].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdatePut2front(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_bToFront);
}



/// <summary>
/// Suggests the keys1.
/// </summary>
void CChildView::suggestKeys1()
{
	m_bLockPrg1 = true;
	m_keyFinder.suggestKeys1();
}

// (OLD suggestKeys1 body moved to KeyFinder::suggestKeys1)



/// <summary>
/// Creates the temporary key arrays1. (Moved to KeyFinder)
/// </summary>
/// <returns></returns>
// int CChildView::createTempKeyArrays1() -- moved to KeyFinder::createTempKeyArrays1()


/// <summary>
/// Suggests the keys2.
/// </summary>
void CChildView::suggestKeys2()
{
	m_bLockPrg2 = true;
	m_keyFinder.suggestKeys2();
}

// (OLD suggestKeys2 body moved to KeyFinder::suggestKeys2)


/// <summary>
/// Creates the temporary key arrays2. (Moved to KeyFinder)
/// </summary>
/// <returns></returns>
// int CChildView::createTempKeyArrays2() -- moved to KeyFinder::createTempKeyArrays2()


/// <summary>
/// Clears the possible keys.
/// </summary>
void CChildView::clearPossibleKeys()
{
	m_keyFinder.clearPossibleKeys();
}



/// <summary>
/// Sorts three numbers
/// </summary>
/// <param name="a">a.</param>
/// <param name="b">The b.</param>
/// <param name="c">The c.</param>
inline void CChildView::sort3(int& a, int& b, int& c)
{
	if (c < b) std::swap(c, b);
	if (b < a) std::swap(b, a);
	if (c < b) std::swap(c, b);
}



/// <summary>
/// Suggests the keys1 thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT SuggestKeys1ThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->suggestKeys1();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Suggests the keys2 thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT SuggestKeys2ThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->suggestKeys2();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Mutuals the check thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT MutualCheckThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->mutualCheck();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Finds the sims thread proc.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT FindSimsThreadProc(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->findSims();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Finds the sims thread proc1.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT FindSimsThreadProc1(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->findSims1();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Finds the sims thread proc2.
/// </summary>
/// <param name="pParam">The p parameter.</param>
/// <returns></returns>
UINT FindSimsThreadProc2(LPVOID pParam)
{
	CChildView* pWnd = static_cast<CChildView*>(pParam);
	pWnd->findSims2();
	AfxEndThread(0);
	return 0;
}



/// <summary>
/// Mutuals the check.
/// </summary>
/// <returns></returns>
bool CChildView::mutualCheck()
{
	m_keyFinder.resetBestKeyComb();
	int tmpRslt = 0;
	// Cascade check
	if (m_pKeyProgressBar1 && m_pKeyProgressBar2)
	{
		PostMessage(CM_UPDATE_KEYPROGRESS1, 0, 0);
		PostMessage(CM_UPDATE_KEYPROGRESS2, 0, 0);
	}
	if (m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, 0) == 0  && m_keyFinder.getNumberOfPossibleKeys(2, SUGKEYS, 0) == 0)
	{
		MessageBox(CMsg(IDS_NTHR_TBL_KEY_FND)); // CMsg(IDS_NTHR_TBL_KEY_FND)
		return false;
	}
	if (m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, 0) == 0)
	{
		MessageBox(CMsg(IDS_NO_KEY_FND_IN_FRST)); // CMsg(IDS_NO_KEY_FND_IN_FRST)
		return false;
	}
	if (m_keyFinder.getNumberOfPossibleKeys(2, SUGKEYS, 0) == 0)
	{
		MessageBox(CMsg(IDS_NO_KEY_FND_IN_SCND)); // CMsg(IDS_NO_KEY_FND_IN_SCND)
		return false;
	}
	int m_i = 0;
	m_bLockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	int order = 1;
	int pkCnt1 = m_keyFinder.getPossibleKeyCounter1();
	while (m_i <= pkCnt1 && tmpRslt < 100)
	{
		prgHlpr0 = 100 * m_i / (pkCnt1 > 0 ? pkCnt1 : 1);
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
			PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
		}
		if (m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, m_i) == order)
		{
			tmpRslt = m_keyFinder.checkKeys(m_i);
		}
		else
		{
			break;
		}
		m_i++;
	}
	order++;
	int maxOrder = m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, (pkCnt1 - 1 >= 0 ? pkCnt1 - 1 : 0));
	while (tmpRslt < 90 && order <= maxOrder)
	{
		while (m_i <= pkCnt1 && tmpRslt < 90)
		{
			prgHlpr0 = 100 * m_i / (pkCnt1 > 0 ? pkCnt1 : 1);
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
			}
			if (m_keyFinder.getNumberOfPossibleKeys(1, order, m_i) == order)
			{
				tmpRslt = m_keyFinder.checkKeys(m_i);
			}
			else
			{
				break;
			}
			m_i++;
		}
		order++;
	}
	if (tmpRslt)
	{
		PostMessage(CM_UPDATE_PROGRESS, 0, 1e5);
		return true;
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 2e5);
	return false;
}



/// <summary>
/// Checks the keys.
/// </summary>
/// <param name="tab1">The tab1.</param>
/// <returns></returns>
int CChildView::checkKeys(int tab1)
{
	m_bLockPrg1 = true;
	return m_keyFinder.checkKeys(tab1);
}

// (OLD checkKeys body moved to KeyFinder::checkKeys)



/// <summary>
/// Deletes the key.
/// </summary>
/// <param name="table">The table.</param>
/// <param name="column">The column.</param>
/// <returns></returns>
int CChildView::deleteKey(int table, int column)
{
	return m_engine.deleteKey(table, column);
}



/// <summary>
/// Sets the key.
/// </summary>
/// <param name="table">The table.</param>
/// <param name="column">The column.</param>
void CChildView::setKey(int table, int column)
{
	m_engine.setNthKey(m_engine.getKeyPairCounter(), table == 1 ? column : 0, table == 2 ? column : 0);
}



/// <summary>
/// Deletes all keys.
/// </summary>
void CChildView::deleteAllKeys()
{
	m_engine.deleteAllKeys();
	m_bToDisplaySimilarClms = false;
	m_bXSimilarityComputed = false;
	m_vecSimilaritiesAcrossTables.clear();
	m_vecSimilaritiesAcrossTablesSorted.clear();
}



/// <summary>
/// Determines whether there any keys.
/// </summary>
/// <returns></returns>
bool CChildView::areThereAnyKeys()
{
	return m_engine.areThereAnyKeys();
}



/// <summary>
/// Determines whether [is this a key] [the specified table].
/// </summary>
/// <param name="table">The table.</param>
/// <param name="column">The column.</param>
/// <returns>
///   <c>true</c> if [is this a key] [the specified table]; otherwise, <c>false</c>.
/// </returns>
bool CChildView::isThisAKey(int table, int column)
{
	return m_engine.isThisAKey(table, column);
}



/// <summary>
/// Gets the key.
/// </summary>
/// <param name="table">The table.</param>
/// <param name="key">The key.</param>
/// <returns></returns>
int CChildView::getNthKey(int table, int key)
{
	return m_engine.getNthKey(table, key);
}



/// <summary>
/// Called when [r button up].
/// </summary>
/// <param name="nFlags">The n flags.</param>
/// <param name="point">The point.</param>
void CChildView::OnRButtonUp(UINT nFlags, CPoint point)
{
	if (M_CCell.x * M_CCell.y)
	{
		if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns)
		{
			if (deleteKey(1, M_CCell.y) + deleteKey(2, M_CCell.x) == 0)
			{
				pushKey(M_CCell.y, M_CCell.x);
			}
			this->Invalidate();
		}
	}
	CWnd::OnRButtonUp(nFlags, point);
}



/// <summary>
/// Sets the NTH key.
/// </summary>
/// <param name="n">The n.</param>
/// <param name="col1">The col1.</param>
/// <param name="col2">The col2.</param>
void CChildView::setNthKey(int n, int col1, int col2)
{
	m_engine.setNthKey(n, col1, col2);
}



/// <summary>
/// Inserts the key at cols.
/// </summary>
/// <param name="n">The n.</param>
/// <param name="col1">The col1.</param>
/// <param name="col2">The col2.</param>
void CChildView::insertKeyAt(int n, int col1, int col2)
{
	m_engine.insertKeyAt(n, col1, col2);
}



/// <summary>
/// Deletes the key at position.
/// </summary>
/// <param name="n">The n.</param>
void CChildView::deleteKeyAt(int n)
{
	m_engine.deleteKeyAt(n);
}



/// <summary>
/// Pushes the key.
/// </summary>
/// <param name="col1">The col1.</param>
/// <param name="col2">The col2.</param>
void CChildView::pushKey(int col1, int col2)
{
	m_engine.pushKey(col1, col2);
}



/// <summary>
/// Uses the possible keys.
/// </summary>
/// <returns></returns>
bool CChildView::usePossibleKeys()
{
	deleteAllKeys();
	BestKeyComb best = m_keyFinder.getBestKeyComb();
	int n = m_keyFinder.getNumberOfPossibleKeys();
	for (int tmp_i = 0; tmp_i <= n; tmp_i++)
	{
		int k1 = m_keyFinder.getPossibleKey1(best.pk1, tmp_i);
		int k2 = m_keyFinder.getPossibleKey2(best.pk2, tmp_i);
		if (k1 + k2)
		{
			pushKey(k1, k2);
		}
	}
	return false;
}



/// <summary>
/// Gets the number of possible keys.
/// </summary>
/// <returns></returns>
int CChildView::getNumberOfPossibleKeys()
{
	return m_keyFinder.getNumberOfPossibleKeys();
}



/// <summary>
/// Sorts the examined keys.
/// </summary>
/// <param name="table">The table.</param>
// void CChildView::sortExaminedKeys() -- moved to KeyFinder
// int  CChildView::sumExaminedKeys()  -- moved to KeyFinder
// bool CChildView::is2BExaminedOnce() -- moved to KeyFinder
// bool CChildView::getSimilarKeyProbability() -- moved to KeyFinder
// int  CChildView::getNthEntropy()    -- moved to KeyFinder
// int  CChildView::CalculateEntropyRank() -- moved to KeyFinder
// bool CChildView::isEntropyStored()  -- moved to KeyFinder


/// <summary>
/// Called when [update combo2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCombo2(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(true);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	if (!m_pCombo2)
	{
		m_pCombo2 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_COMBO2));
		if (m_pCombo2)
		{
			m_pCombo2->SelectItem(1);
		}
	}
}



/// <summary>
/// Called when [combo2].
/// </summary>
void CChildView::OnCombo2()
{
	int complexity = 100000;
	if (m_pCombo2->GetCurSel() == 0) complexity = 10000;
	if (m_pCombo2->GetCurSel() == 1) complexity = 100000;
	if (m_pCombo2->GetCurSel() == 2) complexity = 1000000;
	m_keyFinder.setComplexity(complexity);
}



//int CChildView::getPossibleKeyReadiness(int table, int order)
//{
//
//	return 0;
//}



/// <summary>
/// Gets the number of possible keys.
/// </summary>
/// <param name="table">The table.</param>
/// <param name="order">The order.</param>
/// <param name="item">The item.</param>
/// <returns></returns>
int CChildView::getNumberOfPossibleKeys(int table, int order, int item)
{
	return m_keyFinder.getNumberOfPossibleKeys(table, order, item);
}



/// <summary>
/// Finds the sims.
/// </summary>
void CChildView::findSims() // do not use in case there is a sufficient RAM capacity
{
	long index[2];
	COleVariant vData;
	CString szdata;
	long tmpSim;
	int prgHlpr0, prgHlpr;
	prgHlpr = prgHlpr0 = 0;
	m_vecSimilaritiesAcrossTables.clear();
	m_vecSimilaritiesAcrossTablesSorted.clear();
	SimilaritiesAcrossTables tempSimilarity;
	m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);
	for (int tmp_i = 1; tmp_i <= m_Table1.NumberOfColumns + 1; tmp_i++)
	{
		tempSimilarity.similarityOrder = 0;
		tempSimilarity.similarity = 0;
		tempSimilarity.clm1 = tmp_i;
		tempSimilarity.clm2 = 0;
		m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);;
	}
	for (int c_i1 = 1; c_i1 <= m_Table1.NumberOfColumns; c_i1++)
	{
		prgHlpr0 = 100 * c_i1 / m_Table1.NumberOfColumns; // only 3 keys
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
		}
		m_mapTmpMap1.clear();
		for (int r_i1 = m_Table1.FirstRowWithData; r_i1 <= m_Table1.NumberOfRows; r_i1++)
		{
			szdata = m_excel1.getCellValue(c_i1, r_i1);
			if ((szdata != L"") && (m_mapTmpMap1.find(szdata) == m_mapTmpMap1.end()))
			{
				m_mapTmpMap1[szdata] = r_i1;
			}
		}
		for (int c_i2 = 1; c_i2 <= m_Table2.NumberOfColumns; c_i2++)
		{
			tmpSim = 0;
			m_mapTmpMap2.clear();
			for (int r_i2 = m_Table2.FirstRowWithData; r_i2 <= m_Table2.NumberOfRows; r_i2++)
			{
				szdata = m_excel2.getCellValue(c_i2, r_i2);
				if ((szdata != L"") && (m_mapTmpMap1.find(szdata) != m_mapTmpMap1.end()))
				{		
					if (m_mapTmpMap2.find(szdata) == m_mapTmpMap2.end())
					{
						m_mapTmpMap2[szdata] = r_i2;
						tmpSim++;
					}
				}
			}
			if (tmpSim > m_vecSimilaritiesAcrossTables[c_i1].similarity)
			{
				m_vecSimilaritiesAcrossTables[c_i1].similarity = tmpSim;
				m_vecSimilaritiesAcrossTables[c_i1].clm1 = c_i1;
				m_vecSimilaritiesAcrossTables[c_i1].clm2 = c_i2;
			}
		}
	}
	int simOrder = 1;
	tempSimilarity.clm1 = 0;
	tempSimilarity.clm2 = 0;
	tempSimilarity.similarity = 0;
	tempSimilarity.similarityOrder = 0;
	m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
	for (int i0 = 1; i0 <= m_Table1.NumberOfColumns; i0++)
	{
		tempSimilarity.clm1 = 0;
		tempSimilarity.clm2 = 0;
		tempSimilarity.similarity = 0;
		tempSimilarity.similarityOrder = 0;
		for (int i1 = 1; i1 <= m_Table1.NumberOfColumns; i1++)
		{
			if ( m_vecSimilaritiesAcrossTables[i1].similarity > 0 && m_vecSimilaritiesAcrossTables[i1].similarity > tempSimilarity.similarity && m_vecSimilaritiesAcrossTables[i1].similarityOrder == 0 ) // clm2 only serves here for storing of the actual measured similarity
			{
				tempSimilarity.similarityOrder = simOrder;
				tempSimilarity.similarity = m_vecSimilaritiesAcrossTables[i1].similarity;
				tempSimilarity.clm1 = m_vecSimilaritiesAcrossTables[i1].clm1;
				tempSimilarity.clm2 = m_vecSimilaritiesAcrossTables[i1].clm2;
			}
		}
		if (tempSimilarity.similarity > 0)
		{
			simOrder++;
			m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
			m_vecSimilaritiesAcrossTables[tempSimilarity.clm1].similarityOrder = simOrder;
		}
	}
	m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder = simOrder - 1; // at the zero position, there will be stored the total number of all the columns that have a "lookalike" in the second file
	PostMessage(CM_UPDATE_PROGRESS, 0, 0);
	this->Invalidate();
	if (simOrder > 1)
	{
		m_bToDisplaySimilarClms = true;
		m_bXSimilarityComputed = true;
	}
	m_bLockPrg1 = false;
	return;
}



/// <summary>
/// Finds the sims1.
/// </summary>
void CChildView::findSims1()
{
	long index[2];
	COleVariant vData;
	CString szdata;
	long long tmpSim;
	int prgHlpr0, prgHlpr;
	prgHlpr = prgHlpr0 = 0;
	std::map<CString, long> thdSafe_tmpMap1; // searching for appropriate keys
	std::map<CString, long> thdSafe_tmpMap2;
	//typedef	std::map<CString, long>::iterator Iterator;
	CString what = L"";
	long occurence1 = 0;
	long occurence2 = 0;
	long size1 = 0;
	long size2 = 0;
	long minsize = 0;
	long maxsize = 0;
	double tmpUnitSim = 0.f;
	long tmp_varRat = 0;
	long sim = 0;
	long sumOccurence1 = 0;
	long sumOccurence2 = 0;
	long pureSim;
	int tmp_bnd_hlf = m_Table1.NumberOfColumns / 2;
	for (int c_i1 = 1; c_i1 <= tmp_bnd_hlf; c_i1++)
	{
		prgHlpr0 = 100 * c_i1 / tmp_bnd_hlf; // only 3 keys
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
		}
		thdSafe_tmpMap1.clear();
		for (int r_i1 = m_Table1.FirstRowWithData; r_i1 <= m_Table1.NumberOfRows; r_i1++)
		{
			szdata = m_excel1.getCellValue(c_i1, r_i1);
			if (szdata != L"")
			{
				if (thdSafe_tmpMap1.find(szdata) == thdSafe_tmpMap1.end())
				{
					thdSafe_tmpMap1[szdata] = 1;
				}
				else
				{
					thdSafe_tmpMap1[szdata] = thdSafe_tmpMap1[szdata] + 1;
				}
			}
		}
		for (int c_i2 = 1; c_i2 <= m_Table2.NumberOfColumns; c_i2++)
		{
			thdSafe_tmpMap2.clear();
			for (int r_i2 = m_Table2.FirstRowWithData; r_i2 <= m_Table2.NumberOfRows; r_i2++)
			{
				szdata = m_excel2.getCellValue(c_i2, r_i2);
				if ((szdata != L"") && (thdSafe_tmpMap1.find(szdata) != thdSafe_tmpMap1.end()))
				{
					if (thdSafe_tmpMap2.find(szdata) == thdSafe_tmpMap2.end())
					{
						thdSafe_tmpMap2[szdata] = 1;
						//tmpSim++;
					}
					else
					{
						thdSafe_tmpMap2[szdata] = thdSafe_tmpMap2[szdata] + 1;
					}
				}
			}
			sumOccurence1 = sumOccurence2 = 0;
			tmpSim = 0;
			for (auto iterator: thdSafe_tmpMap1)
			{
				what = iterator.first;
				occurence1 = iterator.second;
				sumOccurence1 += occurence1;
				occurence2 = 0;
				if (thdSafe_tmpMap2.find(what) != thdSafe_tmpMap2.end())
				{
					occurence2 = thdSafe_tmpMap2[what];
					sumOccurence2 += occurence2;
					tmpUnitSim = max(occurence1, occurence2) - min(occurence1, occurence2);
					tmpSim += (tmpUnitSim);
				}
			}
			sim = tmpSim;
			size1 = m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1; //size1 = thdSafe_tmpMap1.size();
			size2 = m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1; //size2 = thdSafe_tmpMap2.size();
			minsize = min(size1, size2);
			minsize = minsize ? minsize : 1;
			maxsize = max(size1, size2);
			{
				tmp_varRat = min(thdSafe_tmpMap1.size(), thdSafe_tmpMap2.size());
				if (tmp_varRat)
				{
					tmp_varRat = tmp_varRat ? tmp_varRat : 1;
					tmp_varRat = (minsize < tmp_varRat ? 1 : minsize / tmp_varRat);
					sim = (minsize - sim) / tmp_varRat + 1;
				}
				if (sim == 0 && thdSafe_tmpMap1.size() == thdSafe_tmpMap2.size() && tmp_varRat > 0)
				{
					sim = 1;
				}
			}
			//if (c_i1 == 16)
			//{
			//	//CString s;
			//	//s.Format(L"c_i1: %n    c_i2: %n    tmp_varRat: %n", c_i1, c_i2, tmp_varRat);
			//	TRACE(L"c_i1: %i    c_i2: %i    thdSafe_tmpMap1.size(): %i     thdSafe_tmpMap2.size(): %i     tmpSim: %i    sumOccurence1: %i    sumOccurence2: %i\n", c_i1, c_i2, thdSafe_tmpMap1.size(), thdSafe_tmpMap2.size(), tmpSim, sumOccurence1, sumOccurence2);
			//}
				pureSim = (maxsize - abs(sumOccurence2 - sumOccurence1)) - tmpSim;
			if (pureSim > m_vecSimilaritiesAcrossTables[c_i1].pureSim && sim > 0)
			{
				m_vecSimilaritiesAcrossTables[c_i1].similarity = min(thdSafe_tmpMap1.size(), thdSafe_tmpMap2.size());
				m_vecSimilaritiesAcrossTables[c_i1].clm1 = c_i1;
				m_vecSimilaritiesAcrossTables[c_i1].clm2 = c_i2;
				m_vecSimilaritiesAcrossTables[c_i1].pureSim = pureSim;
			}
		}
	}
	PostMessage(CM_UPDATE_KEYPROGRESS1, 0, 1000);
	return;
}



/// <summary>
/// Finds the sims2.
/// </summary>
void CChildView::findSims2()
{
	long index[2];
	COleVariant vData;
	CString szdata;
	long long tmpSim;
	int prgHlpr0, prgHlpr;
	prgHlpr = prgHlpr0 = 0;
	std::map<CString, long> thdSafe_tmpMap1; // searching for appropriate keys
	std::map<CString, long> thdSafe_tmpMap2;
	//typedef	std::map<CString, long>::iterator Iterator;
	CString what = L"";
	long occurence1 = 0;
	long occurence2 = 0;
	long size1 = 0;
	long size2 = 0;
	long minsize = 0;
	long maxsize = 0;
	double tmpUnitSim = 0.f;
	long tmp_varRat = 0;
	long sim = 0;
	long sumOccurence1 = 0;
	long sumOccurence2 = 0;
	long pureSim;
	int tmp_bnd_hlf = m_Table1.NumberOfColumns / 2;
	for (int c_i1 = m_Table1.NumberOfColumns - tmp_bnd_hlf; c_i1 <= m_Table1.NumberOfColumns; c_i1++)
	{
		prgHlpr0 = 100 * (c_i1 - tmp_bnd_hlf) / (m_Table1.NumberOfColumns - tmp_bnd_hlf); // only 3 keys
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_KEYPROGRESS2, 0, prgHlpr);
		}
		thdSafe_tmpMap1.clear();
		for (int r_i1 = m_Table1.FirstRowWithData; r_i1 <= m_Table1.NumberOfRows; r_i1++)
		{
			szdata = m_excel1.getTmpCellValue(c_i1, r_i1);
			if (szdata != L"")
			{
				if (thdSafe_tmpMap1.find(szdata) == thdSafe_tmpMap1.end())
				{
					thdSafe_tmpMap1[szdata] = 1;
				}
				else
				{
					thdSafe_tmpMap1[szdata] = thdSafe_tmpMap1[szdata] + 1;
				}
			}
		}
		for (int c_i2 = 1; c_i2 <= m_Table2.NumberOfColumns; c_i2++)
		{
			thdSafe_tmpMap2.clear();
			for (int r_i2 = m_Table2.FirstRowWithData; r_i2 <= m_Table2.NumberOfRows; r_i2++)
			{
				szdata = m_excel2.getTmpCellValue(c_i2, r_i2);
				if ((szdata != L"") && (thdSafe_tmpMap1.find(szdata) != thdSafe_tmpMap1.end()))
				{
					if (thdSafe_tmpMap2.find(szdata) == thdSafe_tmpMap2.end())
					{
						thdSafe_tmpMap2[szdata] = 1;
						//tmpSim++;
					}
					else
					{
						thdSafe_tmpMap2[szdata] = thdSafe_tmpMap2[szdata] + 1;
					}
				}
			}
			sumOccurence1 = sumOccurence2 = 0;
			tmpSim = 0;
			for (auto iterator: thdSafe_tmpMap1)
			{
				what = iterator.first;
				occurence1 = iterator.second;
				sumOccurence1 += occurence1;
				occurence2 = 0;
				if (thdSafe_tmpMap2.find(what) != thdSafe_tmpMap2.end())
				{
					occurence2 = thdSafe_tmpMap2[what];
					sumOccurence2 += occurence2;
					tmpUnitSim = max(occurence1, occurence2) - min(occurence1, occurence2);
					tmpSim += (tmpUnitSim);
				}
			}
			sim = tmpSim;
			size1 = m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1; //size1 = thdSafe_tmpMap1.size();
			size2 = m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1; //size2 = thdSafe_tmpMap2.size();
			minsize = min(size1, size2);
			minsize = minsize ? minsize : 1;
			maxsize = max(size1, size2);
			{
				tmp_varRat = min(thdSafe_tmpMap1.size(), thdSafe_tmpMap2.size());
				if (tmp_varRat)
				{
					tmp_varRat = tmp_varRat ? tmp_varRat : 1;
					tmp_varRat = (minsize < tmp_varRat ? 1 : minsize / tmp_varRat);
					sim = (minsize - sim) / tmp_varRat + 1;
				}
				if (sim == 0 && thdSafe_tmpMap1.size() == thdSafe_tmpMap2.size() && tmp_varRat > 0)
				{
					sim = 1;
				}
			}
			//{
			//	//CString s;
			//	//s.Format(L"c_i1: %n    c_i2: %n    tmp_varRat: %n", c_i1, c_i2, tmp_varRat);
			//	TRACE(L"c_i1: %i    c_i2: %i    thdSafe_tmpMap1.size(): %i     thdSafe_tmpMap2.size(): %i     tmpSim: %i    sumOccurence1: %i    sumOccurence2: %i\n", c_i1, c_i2, thdSafe_tmpMap1.size(), thdSafe_tmpMap2.size(), tmpSim, sumOccurence1, sumOccurence2);
			//}
			pureSim = (maxsize - abs(sumOccurence2 - sumOccurence1)) - tmpSim;
			if (pureSim > m_vecSimilaritiesAcrossTables[c_i1].pureSim && sim > 0)
			{
				m_vecSimilaritiesAcrossTables[c_i1].similarity = min(thdSafe_tmpMap1.size(), thdSafe_tmpMap2.size());
				m_vecSimilaritiesAcrossTables[c_i1].clm1 = c_i1;
				m_vecSimilaritiesAcrossTables[c_i1].clm2 = c_i2;
				m_vecSimilaritiesAcrossTables[c_i1].pureSim = pureSim;
			}
		}
	}
	PostMessage(CM_UPDATE_KEYPROGRESS2, 0, 2000);
	return;
}



/// <summary>
/// Called when [similarpaircheckbox].
/// </summary>
void CChildView::OnSimilarpaircheckbox()
{
	if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns == 0)
	{
		return;
	}
	else
	{
		m_bToDisplaySimilarClms = !m_bToDisplaySimilarClms;
		m_pShowSims->Redraw();
		this->Invalidate();
	}
}



/// <summary>
/// Called when [update similarpaircheckbox].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateSimilarpaircheckbox(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(m_Table1.NumberOfColumns * m_Table2.NumberOfColumns && m_bXSimilarityComputed);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pShowSims = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_SIMILARPAIRCHECKBOX));
	pCmdUI->SetCheck(m_bToDisplaySimilarClms);
}



/// <summary>
/// Called when [findrel BTN].
/// </summary>
void CChildView::OnFindrelBtn()
{
	if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns == 0)
	{
		return;
	}
	if (m_bLockPrg1 || m_bLockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}
	// <Preparation for actual-relatonions check>
	m_vecSimilaritiesAcrossTables.clear();
	m_vecSimilaritiesAcrossTablesSorted.clear();
	SimilaritiesAcrossTables tempSimilarity;
	tempSimilarity.clm1 = 0;
	tempSimilarity.clm2 = 0;
	tempSimilarity.similarity = 0;
	tempSimilarity.similarityOrder = 0;
	m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);
	for (int tmp_i = 1; tmp_i <= m_Table1.NumberOfColumns + 1; tmp_i++)
	{
		tempSimilarity.similarityOrder = 0;
		tempSimilarity.similarity = 0;
		tempSimilarity.clm1 = tmp_i;
		tempSimilarity.clm2 = 0;
		m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);;
	}
	// </Preparation for actual-relations check>
	m_bToDisplaySimilarClms = false;
	m_bXSimilarityComputed = false;
	//AfxBeginThread(FindSimsThreadProc, this);
	AfxBeginThread(FindSimsThreadProc1, this);
	AfxBeginThread(FindSimsThreadProc2, this);
	m_bLockPrg1 = true;
	m_bLockPrg2 = true;
}



/// <summary>
/// Called when [idxcrt BTN].
/// </summary>
void CChildView::OnIdxcrtBtn()
{
}



/// <summary>
/// Called when [cm update key progress1].
/// </summary>
/// <param name="wParam">The w parameter.</param>
/// <param name="lParam">The l parameter.</param>
/// <returns></returns>
afx_msg LRESULT CChildView::OnCmUpdateKeyProgress1(WPARAM wParam, LPARAM lParam)
{
	if ((UINT)lParam > 99)
	{
		m_pKeyProgressBar1->SetPos(0);
		if ((UINT)lParam == 1000)
		{
			m_bLockPrg1 = false;
			if (m_bLockPrg2 == false)
			{
				finishFindRelations();
			}
		}
	}
	else
	{
		m_pKeyProgressBar1->SetPos((UINT)lParam);
	}
	return 0;
}



/// <summary>
/// Called when [cm update key progress2].
/// </summary>
/// <param name="wParam">The w parameter.</param>
/// <param name="lParam">The l parameter.</param>
/// <returns></returns>
afx_msg LRESULT CChildView::OnCmUpdateKeyProgress2(WPARAM wParam, LPARAM lParam)
{
	if ((UINT)lParam > 99)
	{
		m_pKeyProgressBar2->SetPos(0);
		if ((UINT)lParam == 2000)
		{
			m_bLockPrg2 = false;
			if (m_bLockPrg1 == false)
			{
				finishFindRelations();
			}
		}
	}
	else
	{
		m_pKeyProgressBar2->SetPos((UINT)lParam);
	}
	return 0;
}



/// <summary>
/// Called when [update key progress1].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateKeyProgress1(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pKeyProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_KEY_PROGRESS1));
}



/// <summary>
/// Called when [update key progress2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateKeyProgress2(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pKeyProgressBar2 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_KEY_PROGRESS2));
}



/// <summary>
/// Finishes the find relations.
/// </summary>
void CChildView::finishFindRelations()
{
	if (m_bXSimilarityComputed)
	{
		return;
	}
	SimilaritiesAcrossTables tempSimilarity;
	int simOrder = 1;
	tempSimilarity.clm1 = 0;
	tempSimilarity.clm2 = 0;
	tempSimilarity.similarity = 0;
	tempSimilarity.similarityOrder = 0;
	m_vecSimilaritiesAcrossTablesSorted.clear();
	m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
	for (int i0 = 1; i0 <= m_Table1.NumberOfColumns; i0++)
	{
		tempSimilarity.clm1 = 0;
		tempSimilarity.clm2 = 0;
		tempSimilarity.similarity = -1;
		tempSimilarity.similarityOrder = 0;
		for (int i1 = 1; i1 <= m_Table1.NumberOfColumns; i1++)
		{
			if (/*similaritiesAcrossTables[i1].similarity > 0 && */ m_vecSimilaritiesAcrossTables[i1].similarity > tempSimilarity.similarity && m_vecSimilaritiesAcrossTables[i1].similarityOrder == 0) // clm2 only serves here for storing of the actual measured similarity
			{
				tempSimilarity.similarityOrder = simOrder;
				tempSimilarity.similarity = m_vecSimilaritiesAcrossTables[i1].similarity;
				tempSimilarity.clm1 = m_vecSimilaritiesAcrossTables[i1].clm1;
				tempSimilarity.clm2 = m_vecSimilaritiesAcrossTables[i1].clm2;
			}
		}
		/*if (tempSimilarity.similarity > 0)*/
		{
			simOrder++;
			m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
			m_vecSimilaritiesAcrossTables[tempSimilarity.clm1].similarityOrder = simOrder;
		}
	}
	m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder = simOrder - 1; // at the zero position, there will be stored the total number of all the columns that have a "lookalike" in the second file
	this->Invalidate();
	if (simOrder > 1)
	{
		m_bToDisplaySimilarClms = true;
		m_bXSimilarityComputed = true;
	}
}



/// <summary>
/// Called when [update index checkbox].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateIdxCheckbox(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_IDX_CHECKBOX));
	pCmdUI->SetCheck(m_bUseIndexes);
}



/// <summary>
/// Reverses the find.
/// </summary>
/// <param name="lpszData">The LPSZ data.</param>
/// <param name="lpszSub">The LPSZ sub.</param>
/// <param name="startpos">The startpos.</param>
/// <returns></returns>
int CChildView::ReverseFind(LPCTSTR lpszData, LPCTSTR lpszSub, int startpos)
{
	int lenSub = lstrlen(lpszSub);
	int len = lstrlen(lpszData);
	if (0 < lenSub && 0 < len)
	{
		if (startpos == -1 || startpos >= len) startpos = len - 1;
		for (LPCTSTR lpszReverse = lpszData + startpos;
			lpszReverse != lpszData; --lpszReverse)
			if (_tcsncmp(lpszSub, lpszReverse, lenSub) == 0)
				return (lpszReverse - lpszData);
	}
	return -1;
}



/// <summary>
/// Called when [check index].
/// </summary>
void CChildView::OnCheckIdx()
{
	m_bUseIndexes = !m_bUseIndexes;
	m_engine.m_bUseIndexes = m_bUseIndexes;
}



/// <summary>
/// Called when [update check index].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCheckIdx(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_IDX_CHECKBOX));
	pCmdUI->SetCheck(m_bUseIndexes);
}



/// <summary>
/// Called when [usidx check].
/// </summary>
void CChildView::OnUsidxCheck()
{
	m_bUseIndexes = !m_bUseIndexes;
	m_engine.m_bUseIndexes = m_bUseIndexes;
	if (m_bUseIndexes) MessageBox(CMsg(IDS_IDXING_WARNING)); // CMsg(IDS_IDXING_WARNING)
}



/// <summary>
/// Called when [update usidx check].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateUsidxCheck(CCmdUI* pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_USIDX_CHECK));
	pCmdUI->SetCheck(m_bUseIndexes);
}



/// <summary>
/// Called when [update rows1].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateRows1(CCmdUI* pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pRows1 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_ROWS1));
}



/// <summary>
/// Called when [rows1].
/// </summary>
void CChildView::OnRows1()
{
	long prevVal = m_Table1.NumberOfRows;
	CString tmpo = m_pRows1->GetEditText();
	tmpo.Remove(160);
	long tmpi = _ttoi(tmpo);
	m_Table1.NumberOfRows = tmpi > m_Table1.MaxNumberOfRows ? m_Table1.MaxNumberOfRows : tmpi;
	CString tmps;
	tmps.Format(L"%i", m_Table1.NumberOfRows);
	tmps.Remove(160);
	if (prevVal != tmpi)
	{
		m_pRows1->SetEditText(tmps);
	}
}



/// <summary>
/// Called when [update cols1].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCols1(CCmdUI* pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pCols1 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_COLS1));
}



/// <summary>
/// Called when [cols1].
/// </summary>
void CChildView::OnCols1()
{
	long prevVal = m_Table2.NumberOfColumns;
	CString tmpo = m_pCols1->GetEditText();
	tmpo.Remove(160);
	long tmpi = _ttoi(tmpo);
	m_Table1.NumberOfColumns = tmpi > m_Table1.MaxNumberOfCols ? m_Table1.MaxNumberOfCols : tmpi;
	CString tmps;
	tmps.Format(L"%i", m_Table1.NumberOfColumns);
	tmps.Remove(160);
	if (prevVal != tmpi)
	{
		m_pCols1->SetEditText(tmps);
	}
}



/// <summary>
/// Called when [update rows2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateRows2(CCmdUI* pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pRows2 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_ROWS2));
}



/// <summary>
/// Called when [rows2].
/// </summary>
void CChildView::OnRows2()
{
	long prevVal = m_Table2.NumberOfRows;
	CString tmpo = m_pRows2->GetEditText();
	tmpo.Remove(160);
	long tmpi = _ttoi(tmpo);
	m_Table2.NumberOfRows = tmpi > m_Table2.MaxNumberOfRows ? m_Table2.MaxNumberOfRows : tmpi;
	CString tmps;
	tmps.Format(L"%i", m_Table2.NumberOfRows);
	tmps.Remove(160);
	if (prevVal != tmpi)
	{
		m_pRows2->SetEditText(tmps);
	}
}



/// <summary>
/// Called when [update cols2].
/// </summary>
/// <param name="pCmdUI">The p command UI.</param>
void CChildView::OnUpdateCols2(CCmdUI* pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pCols2 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_COLS2));
}



/// <summary>
/// Called when [cols2].
/// </summary>
void CChildView::OnCols2()
{
	long prevVal = m_Table2.NumberOfColumns;
	CString tmpo = m_pCols2->GetEditText();
	tmpo.Remove(160);
	long tmpi = _ttoi(tmpo);
	m_Table2.NumberOfColumns = tmpi > m_Table2.MaxNumberOfCols ? m_Table2.MaxNumberOfCols : tmpi;
	CString tmps;
	tmps.Format(L"%i", m_Table2.NumberOfColumns);
	tmps.Remove(160);
	if (prevVal != tmpi)
	{
		m_pCols2->SetEditText(tmps);
	}
}
