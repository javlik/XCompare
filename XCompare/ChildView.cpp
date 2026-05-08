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
    m_Palette[0] = {0, 0, 0};
    m_Palette[1] = {128, 0, 0};
    m_Palette[2] = {0, 128, 0};
    m_Palette[3] = {128, 128, 0};
    m_Palette[4] = {0, 0, 128};
    m_Palette[5] = {128, 0, 128};
    m_Palette[6] = {0, 128, 128};
    m_Palette[7] = {192, 192, 192};
    m_Palette[8] = {192, 220, 192};
    m_Palette[9] = {166, 202, 240};
    m_Palette[10] = {255, 251, 240};
    m_Palette[11] = {160, 160, 164};
    m_Palette[12] = {128, 128, 128};
    m_Palette[13] = {255, 0, 0};
    m_Palette[14] = {0, 255, 0};
    m_Palette[15] = {255, 255, 0};
    m_Palette[16] = {0, 0, 255};
    m_Palette[17] = {255, 0, 255};
    m_Palette[18] = {0, 255, 255};
    m_Palette[19] = {255, 255, 255};
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
    // ExcelConnector instances (m_excel1, m_excel2) initialise themselves in their own constructors
    // m_nCheckedKeys1/2, entropy and possible-keys state are now inside m_keyFinder
    // Members that were zero-initialised as globals but not yet in constructor
    m_bWaitingForKeys = false;
    m_bKeys1done = false;
    m_bKeys2done = false;
    m_bKeysGathering1done = false;
    m_bKeysGathering2done = false;
    // m_nKeyPairCounter is now inside m_engine
    m_nOldx = 0;
    m_nOldy = 0;
    m_Clnt.w = 0;
    m_Clnt.h = 0;
    m_nMatrixDone = 0;
    m_nPrereqDone = 0;
    m_bMarkIdentCols = false;
    m_nCellWidth = 0;
    m_nCellHeight = 0;
    m_nRibbonWidth = 0;
    m_nViewWidth = 0;
    m_nViewHeight = 0;
    m_nHScrollPos = 0;
    m_nVScrollPos = 0;
    m_nHPageSize = 0;
    m_nVPageSize = 0;
    m_nScrolled_X = 0;
    m_nScrolled_Y = 0;
    m_CClickedCell.x = 0;
    m_CClickedCell.y = 0;
    m_CPrevClickedCell.x = 0;
    m_CPrevClickedCell.y = 0;
    m_pRibbon = nullptr;
    m_pProgressBar1 = nullptr;
    m_pProgressBar2 = nullptr;
    m_pKeyProgressBar1 = nullptr;
    m_pKeyProgressBar2 = nullptr;
    m_pCombo2 = nullptr;
    m_pSheetCombo1 = nullptr;
    m_pSheetCombo2 = nullptr;
    m_pSpinner1_Fdata = nullptr;
    m_pSpinner1_Names = nullptr;
    m_pSpinner2_Fdata = nullptr;
    m_pSpinner2_Names = nullptr;
    m_pMarkIn1 = nullptr;
    m_pMarkIn2 = nullptr;
    m_pSlider = nullptr;
    m_pUnhideExcel = nullptr;
    m_pVerifyKeys = nullptr;
    m_pSameNames = nullptr;
    m_pColorPicker1 = nullptr;
    m_pColorPicker2 = nullptr;
    m_pAuto = nullptr;
    m_pFoundDifferences = nullptr;
    m_pLabel0 = nullptr;
    m_pLabel1 = nullptr;
    m_pLabel2 = nullptr;
    m_pToFront = nullptr;
    m_pShowSims = nullptr;
    m_pCreateNewKeys = nullptr;
    m_pButton2 = nullptr;
    m_pUseIndices = nullptr;
    m_pRows1 = nullptr;
    m_pCols1 = nullptr;
    m_pRows2 = nullptr;
    m_pCols2 = nullptr;
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
    M_CCell.x = 0;
    M_CCell.y = 0;
    m_OldCell.x = 0;
    m_OldCell.y = 0;
    m_Table1.NumberOfColumns = 0;
    m_Table2.NumberOfColumns = 0;
    for (int i = 0; i < 256; i++)
    {
        m_SimsPens[i].CreatePen(PS_ENDCAP_FLAT, static_cast<int>(i / 32 + 0.5),
                                RGB(static_cast<int>((255 - i) / 1.5 + 40), static_cast<int>((255 - i) / 1.5 + 40),
                                    static_cast<int>((255 - i) / 1.5 + 40)));
    }
    m_KeyCurvePen.CreatePen(PS_ENDCAP_FLAT, 2, RGB(100, 150, 250));
    m_bUseIndexes = false;
    m_engine.m_bUseIndexes = false;
    m_bNewFile1 = false;
    m_bNewFile2 = false;
    // m_nComplexity is now inside m_keyFinder (default 100000)
}

CChildView::~CChildView()
{
    // Dynamic arrays (std::vector) release memory automatically.
}

BEGIN_MESSAGE_MAP(CChildView, CWnd)
// --- Standard window messages ---
ON_WM_CREATE()
ON_WM_PAINT()
ON_WM_SIZE()
ON_WM_VSCROLL()
ON_WM_HSCROLL()
ON_WM_MOUSEWHEEL()
ON_WM_MOUSEMOVE()
ON_WM_LBUTTONDBLCLK()
ON_WM_LBUTTONUP()
ON_WM_RBUTTONUP()

// --- File and sheet selection ---
ON_COMMAND(ID_PICK_FIRST_FILE, &CChildView::OnPickFirstFile)
ON_COMMAND(ID_PICK_SECOND_FILE, &CChildView::OnPickSecondFile)
ON_UPDATE_COMMAND_UI(IDC_FILENAME1, &CChildView::OnUpdateFilename1)
ON_UPDATE_COMMAND_UI(IDC_FILENAME2, &CChildView::OnUpdateFilename2)
ON_COMMAND(ID_PICK_FIRST_SHEET, &CChildView::OnPickFirstSheet)
ON_UPDATE_COMMAND_UI(ID_PICK_FIRST_SHEET, &CChildView::OnUpdatePickFirstSheet)
ON_COMMAND(ID_PICK_SECOND_SHEET, &CChildView::OnPickSecondSheet)
ON_UPDATE_COMMAND_UI(ID_PICK_SECOND_SHEET, &CChildView::OnUpdatePickSecondSheet)

// --- Table range controls (header row / first data row) ---
ON_COMMAND(ID_SPIN1_NAMES, &CChildView::OnSpin1Names)
ON_UPDATE_COMMAND_UI(ID_SPIN1_NAMES, &CChildView::OnUpdateSpin1Names)
ON_COMMAND(ID_SPIN1_FDATA, &CChildView::OnSpin1Fdata)
ON_UPDATE_COMMAND_UI(ID_SPIN1_FDATA, &CChildView::OnUpdateSpin1Fdata)
ON_COMMAND(ID_SPIN2_NAMES, &CChildView::OnSpin2Names)
ON_UPDATE_COMMAND_UI(ID_SPIN2_NAMES, &CChildView::OnUpdateSpin2Names)
ON_COMMAND(ID_SPIN2_FDATA, &CChildView::OnSpin2Fdata)
ON_UPDATE_COMMAND_UI(ID_SPIN2_FDATA, &CChildView::OnUpdateSpin2Fdata)
ON_COMMAND(ID_ROWS1, &CChildView::OnRows1)
ON_UPDATE_COMMAND_UI(ID_ROWS1, &CChildView::OnUpdateRows1)
ON_COMMAND(ID_COLS1, &CChildView::OnCols1)
ON_UPDATE_COMMAND_UI(ID_COLS1, &CChildView::OnUpdateCols1)
ON_COMMAND(ID_ROWS2, &CChildView::OnRows2)
ON_UPDATE_COMMAND_UI(ID_ROWS2, &CChildView::OnUpdateRows2)
ON_COMMAND(ID_COLS2, &CChildView::OnCols2)
ON_UPDATE_COMMAND_UI(ID_COLS2, &CChildView::OnUpdateCols2)

// --- Comparison matrix ---
ON_COMMAND(ID_CREATE_MATRIX, &CChildView::OnCreateMatrix)
ON_UPDATE_COMMAND_UI(ID_CREATE_MATRIX, &CChildView::OnUpdateCreateMatrix)
ON_COMMAND(ID_SIMILARITY_THRESHOLD, &CChildView::OnSimilarityThreshold)
ON_UPDATE_COMMAND_UI(ID_SIMILARITY_THRESHOLD, &CChildView::OnUpdateSimilarityThreshold)
ON_COMMAND(ID_GOTO_DIFF_IN_FILE1, &CChildView::OnGotoDiffInFile1)
ON_COMMAND(ID_KEY_SEARCH_COMPLEXITY, &CChildView::OnKeySearchComplexity)
ON_UPDATE_COMMAND_UI(ID_KEY_SEARCH_COMPLEXITY, &CChildView::OnUpdateKeySearchComplexity)
ON_COMMAND(ID_BRING_EXCEL_TO_FRONT, &CChildView::OnBringExcelToFront)
ON_UPDATE_COMMAND_UI(ID_BRING_EXCEL_TO_FRONT, &CChildView::OnUpdateBringExcelToFront)

// --- Key finding and column-relation detection ---
ON_COMMAND(ID_FIND_COLUMN_RELATIONS, &CChildView::OnFindColumnRelations)
ON_COMMAND(ID_IDXCRT_BTN, &CChildView::OnIdxcrtBtn)
ON_COMMAND(ID_SHOW_SIMILAR_COLUMNS, &CChildView::OnShowSimilarColumns)
ON_UPDATE_COMMAND_UI(ID_SHOW_SIMILAR_COLUMNS, &CChildView::OnUpdateShowSimilarColumns)
ON_COMMAND(ID_USE_KEY_INDEXING, &CChildView::OnUseKeyIndexing)
ON_UPDATE_COMMAND_UI(ID_USE_KEY_INDEXING, &CChildView::OnUpdateUseKeyIndexing)

// --- Marking and difference display ---
ON_COMMAND(ID_SUGGEST_KEYS, &CChildView::OnSuggestKeys)
ON_UPDATE_COMMAND_UI(ID_SUGGEST_KEYS, &CChildView::OnUpdateSuggestKeys)
ON_COMMAND(ID_COLOR_PICKER2, &CChildView::OnColorPicker2)
ON_UPDATE_COMMAND_UI(ID_COLOR_PICKER2, &CChildView::OnUpdateColorPicker2)
ON_COMMAND(ID_COLOR_PICKER1, &CChildView::OnColorPicker1)
ON_UPDATE_COMMAND_UI(ID_COLOR_PICKER1, &CChildView::OnUpdateColorPicker1)
ON_COMMAND(ID_GOTO_DIFF_IN_FILE2, &CChildView::OnGotoDiffInFile2)
ON_COMMAND(ID_VERIFY_KEYS, &CChildView::OnVerifyKeys)
ON_UPDATE_COMMAND_UI(ID_VERIFY_KEYS, &CChildView::OnUpdateVerifyKeys)
ON_COMMAND(ID_AUTO_MARK, &CChildView::OnAutoMark)
ON_UPDATE_COMMAND_UI(ID_AUTO_MARK, &CChildView::OnUpdateAutoMark)
ON_COMMAND(ID_MARK_IN_FILE1, &CChildView::OnMarkInFile1)
ON_UPDATE_COMMAND_UI(ID_MARK_IN_FILE1, &CChildView::OnUpdateMarkInFile1)
ON_COMMAND(ID_MARK_IN_FILE2, &CChildView::OnMarkInFile2)
ON_UPDATE_COMMAND_UI(ID_MARK_IN_FILE2, &CChildView::OnUpdateMarkInFile2)
ON_COMMAND(ID_SAME_NAMES_ONLY, &CChildView::OnSameNamesOnly)
ON_UPDATE_COMMAND_UI(ID_SAME_NAMES_ONLY, &CChildView::OnUpdateSameNamesOnly)
ON_COMMAND(ID_DIFFS_LIST, &CChildView::OnDiffslist)
ON_UPDATE_COMMAND_UI(ID_DIFFS_LIST, &CChildView::OnUpdateDiffslist)

// --- Display-only controls (progress bars, status labels) ---
ON_UPDATE_COMMAND_UI(ID_PROGRESS1, &CChildView::OnUpdateProgress1)
ON_UPDATE_COMMAND_UI(ID_PROGRESS2, &CChildView::OnUpdateProgress1)
ON_UPDATE_COMMAND_UI(ID_PROGRESS3, &CChildView::OnUpdateProgress2)
ON_UPDATE_COMMAND_UI(ID_KEY_PROGRESS1, &CChildView::OnUpdateKeyProgress1)
ON_UPDATE_COMMAND_UI(ID_KEY_PROGRESS2, &CChildView::OnUpdateKeyProgress2)

// --- Custom window messages: progress updates (lParam = 0–100) ---
ON_MESSAGE(CM_UPDATE_PROGRESS, &CChildView::OnCmUpdateProgress)
ON_MESSAGE(CM_UPDATE_PROGRESS2, &CChildView::OnCmUpdateProgress2)
ON_MESSAGE(CM_UPDATE_PROGRESS3, &CChildView::OnCmUpdateProgress3)
ON_MESSAGE(CM_UPDATE_KEYPROGRESS1, &CChildView::OnCmUpdateKeyProgress1)
ON_MESSAGE(CM_UPDATE_KEYPROGRESS2, &CChildView::OnCmUpdateKeyProgress2)

// --- Custom window messages: worker-thread completion events ---
ON_MESSAGE(CM_KEYS1_DONE, &CChildView::OnCmKeys1Done)
ON_MESSAGE(CM_KEYS2_DONE, &CChildView::OnCmKeys2Done)
ON_MESSAGE(CM_GATHERING1_DONE, &CChildView::OnCmGathering1Done)
ON_MESSAGE(CM_GATHERING2_DONE, &CChildView::OnCmGathering2Done)
ON_MESSAGE(CM_KEYS_FOUND, &CChildView::OnCmKeysFound)
ON_MESSAGE(CM_KEYS_NOT_FOUND, &CChildView::OnCmKeysNotFound)
ON_MESSAGE(CM_FIRSTPASS_DONE, &CChildView::OnCmFirstPassDone)
ON_MESSAGE(CM_MARKING_READY, &CChildView::OnCmMarkingReady)
ON_MESSAGE(CM_SIMS1_DONE, &CChildView::OnCmSims1Done)
ON_MESSAGE(CM_SIMS2_DONE, &CChildView::OnCmSims2Done)
END_MESSAGE_MAP()

BOOL CChildView::PreCreateWindow(CREATESTRUCT& cs)
{
    if (!CWnd::PreCreateWindow(cs))
        return FALSE;
    cs.dwExStyle |= WS_EX_CLIENTEDGE;
    cs.style &= ~WS_BORDER;
    cs.style |= WS_VSCROLL | WS_HSCROLL;
    cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS, ::LoadCursor(NULL, IDC_ARROW),
                                       reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1), NULL);
    m_pRibbon = NULL;
    SYSTEM_INFO sysinfo;
    GetSystemInfo(&sysinfo);
    return TRUE;
}

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
    font1.CreateFontW(16, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                      DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font2.CreateFontW(16, 0, 900, 900, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                      DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font3.CreateFontW(12, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                      DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font4.CreateFontW(30, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                      DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font1B.CreateFontW(16, 0, 0, 0, FW_EXTRABOLD, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS,
                       CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font2B.CreateFontW(16, 0, 900, 900, FW_EXTRABOLD, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS,
                       CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font1C.CreateFontW(12, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                       DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    font2C.CreateFontW(12, 0, 900, 900, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                       DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    PaintCtx ctx;
    ctx.pen1 = &pen1;
    ctx.pen2 = &pen2;
    ctx.pen3 = &pen3;
    ctx.pen4 = &pen4;
    ctx.pen5 = &pen5;
    ctx.pen6 = &pen6;
    ctx.pen7 = &pen7;
    ctx.pen8 = &pen8;
    ctx.pen9 = &pen9;
    ctx.pen10 = &pen10;
    ctx.pen11 = &pen11;
    ctx.pen12 = &pen12;
    ctx.brush0 = &brush0;
    ctx.brush1 = &brush1;
    ctx.brush2 = &brush2;
    ctx.brush3 = &brush3;
    ctx.brush4 = &brush4;
    ctx.brush5 = &brush5;
    ctx.brush6 = &brush6;
    ctx.brush7 = &brush7;
    ctx.font1 = &font1;
    ctx.font2 = &font2;
    ctx.font3 = &font3;
    ctx.font4 = &font4;
    ctx.font1B = &font1B;
    ctx.font2B = &font2B;
    ctx.font1C = &font1C;
    ctx.font2C = &font2C;
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
    // Sets font and draws comparison statistics in the top-left corner of the view.
    dc.SelectObject(ctx.font4);
    CString textPercentage = L""; // Renamed from prcnt to textPercentage
    
    // Display individual cell statistics if similarity view is disabled and a valid cell is selected
    if (!m_bToDisplaySimilarClms && M_CCell.x * M_CCell.y && m_nMatrixDone)
    {
        dc.SetBkMode(TRANSPARENT);
        if (M_CCell.x <= ctx.bnd_X_max && M_CCell.y <= ctx.bnd_Y_max)
        {
            // Number of exactly matching rows between the two columns
            long matchedRowsCount = m_matrix.get(M_CCell.x, M_CCell.y); // Renamed from sameness
            if (m_Table1.Columns[M_CCell.y] == m_Table2.Columns[M_CCell.x] && matchedRowsCount < m_nEffMax)
                dc.SetTextColor(RGB(255, 0, 0)); // Highlight in red if names match but rows differ
            else
                dc.SetTextColor(RGB(0, 0, 0)); // Standard black text
            
            // Print the count of different rows (Delta)
            textPercentage.Format(L"\u0394:%i", m_nEffMax - matchedRowsCount);
            dc.TextOutW(5, 20, textPercentage);
            
            // Print the count of matching rows
            dc.SetTextColor(RGB(0, 255, 0));
            textPercentage.Format(L"=:%i", matchedRowsCount);
            dc.TextOutW(5, 50, textPercentage);
            
            // Mark if both columns are empty
            if (m_engine.isEmptyCol1(M_CCell.y) && m_engine.isEmptyCol2(M_CCell.x))
            {
                dc.SetTextColor(RGB(0, 0, 0));
                dc.TextOutW(5, 80, CMsg(IDS_EMPTY));
            }
        }
    }
    
    // Draw the file names in the top-left corner
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
    
    // Draw similarity values if the similarity overlay is enabled
    if (M_CCell.x * M_CCell.y && m_bToDisplaySimilarClms)
    {
        if (M_CCell.x <= ctx.bnd_X_max && M_CCell.y <= ctx.bnd_Y_max)
        {
            dc.SetTextColor(RGB(50, 100, 250));
            dc.SetBkMode(TRANSPARENT);
            dc.SelectObject(ctx.font1C);
            dc.TextOutW(5, 30, CMsg(IDS_KEY_SUITABILITY));
            dc.SelectObject(ctx.font4);
            textPercentage.Format(L"~ %i%%", 100 * m_vecSimilaritiesAcrossTables[M_CCell.y].similarity /
                                        min(m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1,
                                            m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1));
            dc.TextOutW(15, 60, textPercentage);
        }
    }
}

void CChildView::paintGridLines(CDC& dc, PaintCtx& ctx)
{
    // Prepare the background brushes and pens for the grid
    dc.SelectObject(ctx.brush0);
    dc.SelectObject(ctx.pen2);
    dc.SelectObject(ctx.brush1);
    
    // Draw vertical and horizontal grid lines to separate individual cells
    if (!m_bToDisplaySimilarClms)
    {
        // Draw vertical lines
        for (int coordX = OFFSET_X + STEP_X; coordX < m_Clnt.w; coordX += STEP_X)
        {
            dc.MoveTo(coordX, 0);
            dc.LineTo(coordX, m_Clnt.h);
        }
        // Draw horizontal lines
        for (int coordY = OFFSET_Y + STEP_Y; coordY < m_Clnt.h; coordY += STEP_X)
        {
            dc.MoveTo(0, coordY);
            dc.LineTo(m_Clnt.w, coordY);
        }
    }
    dc.SelectObject(ctx.pen2);
    dc.SelectObject(ctx.brush1);
}

void CChildView::paintRowHeaders(CDC& dc, PaintCtx& ctx)
{
    // Iterate through visible rows on the screen to draw the left column headers (Table 1 columns)
    int adjustedRowIndex; // Actual data row index accounting for scrolling (renamed from mx_y_adj)
    for (int yIndex = ctx.bnd_Y_min; yIndex <= ctx.bnd_Y_max; yIndex++)
    {
        bool isHovered = false; // Indicates if the cursor is hovering over the cell
        adjustedRowIndex = yIndex + m_VisTopLeft.top;
        dc.SetBkMode(OPAQUE);
        
        // Highlight rows that are selected as primary keys
        if (m_engine.isThisAKey(1, adjustedRowIndex))
        {
            if (adjustedRowIndex == m_OldCell.y)
            {
                dc.SelectObject(ctx.brush6);
            }
            else
            {
                if (adjustedRowIndex == M_CCell.y)
                {
                    if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns &&
                        (M_CCell.x > 0 || m_bToDisplaySimilarClms) && M_CCell.x <= m_Table2.NumberOfColumns)
                    {
                        isHovered = true;
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
        
        // Final evaluation of hover state for drawing the hover boundary
        if (adjustedRowIndex == M_CCell.y)
        {
            if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && (M_CCell.x > 0 || m_bToDisplaySimilarClms) &&
                M_CCell.x <= m_Table2.NumberOfColumns)
            {
                isHovered = true;
            }
            else
            {
                if (m_engine.isThisAKey(1, adjustedRowIndex))
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
		      
        // Draw the background rectangle of the header
        dc.Rectangle(0, OFFSET_Y + yIndex * STEP_Y, 1 + OFFSET_X + STEP_X, 1 + OFFSET_Y + yIndex * STEP_Y + STEP_Y);
        
        // Draw the red selection rectangle if hovered
        if (isHovered)
        {
            dc.SetBkMode(TRANSPARENT);
            dc.SelectObject(ctx.brush0);
            if (m_bToDisplaySimilarClms)
                dc.SelectObject(ctx.pen12); // Blueish border in similarity mode
            else
                dc.SelectObject(ctx.pen11); // Red border in default mode
            dc.Rectangle(2, 2 + OFFSET_Y + yIndex * STEP_Y, OFFSET_X + STEP_X - 1,
                         -1 + OFFSET_Y + yIndex * STEP_Y + STEP_Y);
            dc.SetBkMode(OPAQUE);
            dc.SelectObject(ctx.brush0);
            dc.SelectObject(ctx.pen4);
        }
        
        // Draw completely matched column indicator (green dot) or empty column indicator (yellow dot)
        if (m_nMatrixDone && !m_bOnlyPcnt && ((yIndex - m_VisTopLeft.top) > 0))
        {
            if (m_pbGreenClms1[yIndex])
            {
                dc.SelectObject(ctx.pen5);
                dc.SelectObject(ctx.brush1);
                dc.Ellipse(OFFSET_X, OFFSET_X + (yIndex - m_VisTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1,
                           OFFSET_Y + STEP_Y + (yIndex - m_VisTopLeft.top) * STEP_Y);
            }
            if (m_engine.isEmptyCol1(yIndex))
            {
                dc.SelectObject(ctx.pen6);
                dc.SelectObject(ctx.brush2);
                dc.Ellipse(OFFSET_X, OFFSET_X + (yIndex - m_VisTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1,
                           OFFSET_Y + STEP_Y + (yIndex - m_VisTopLeft.top) * STEP_Y);
            }
        }
        
        // Draw the name of the column text
        dc.SetBkMode(TRANSPARENT);
        if (m_engine.isThisAKey(1, adjustedRowIndex))
        {
            dc.SelectObject(ctx.font1B); // Bold if key
            dc.SetTextColor(RGB(0, 0, 170));
        }
        else
        {
            dc.SelectObject(ctx.font1);
            dc.SetTextColor(RGB(0, 0, 0));
        }
        dc.TextOutW(2, OFFSET_Y + 5 + yIndex * STEP_Y, m_Table1.Columns[adjustedRowIndex]);
    }
}

void CChildView::paintColumnHeaders(CDC& dc, PaintCtx& ctx)
{
    // Iterate through visible columns on the screen to draw the top column headers (Table 2 columns)
    int adjustedColIndex; // Actual data column index accounting for scrolling (renamed from mx_x_adj)
    for (int xIndex = ctx.bnd_X_min; xIndex <= ctx.bnd_X_max; xIndex++)
    {
        bool isHovered = false; // Indicates if the cursor is hovering over the cell
        adjustedColIndex = xIndex + m_VisTopLeft.left;
        dc.SetBkMode(OPAQUE);
        
        // Highlight columns that are selected as primary keys
        if (m_engine.isThisAKey(2, adjustedColIndex))
        {
            if (adjustedColIndex == m_OldCell.x)
            {
                dc.SelectObject(ctx.brush6);
            }
            else
            {
                if (adjustedColIndex == M_CCell.x)
                {
                    if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns &&
                        (M_CCell.x > 0 || m_bToDisplaySimilarClms) && M_CCell.x <= m_Table2.NumberOfColumns)
                    {
                        isHovered = true;
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
        
        // Final evaluation of hover state for drawing the hover boundary
        if (adjustedColIndex == M_CCell.x)
        {
            if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && (M_CCell.x > 0 || m_bToDisplaySimilarClms) &&
                M_CCell.x <= m_Table2.NumberOfColumns)
            {
                isHovered = true;
            }
            else
            {
                if (m_engine.isThisAKey(2, adjustedColIndex))
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
        
        // Draw the background rectangle of the header
        dc.Rectangle(OFFSET_X + xIndex * STEP_X, 0, 1 + OFFSET_X + xIndex * STEP_X + STEP_X, 1 + OFFSET_Y + STEP_Y);
        
        // Draw the red selection rectangle if hovered
        if (isHovered)
        {
            dc.SetBkMode(TRANSPARENT);
            dc.SelectObject(ctx.brush0);
            if (m_bToDisplaySimilarClms)
                dc.SelectObject(ctx.pen12); // Blueish border in similarity mode
            else
                dc.SelectObject(ctx.pen11); // Red border in default mode
            dc.Rectangle(2 + OFFSET_X + xIndex * STEP_X, 2, -1 + OFFSET_X + xIndex * STEP_X + STEP_X,
                         OFFSET_Y + STEP_Y - 1);
            dc.SetBkMode(OPAQUE);
            dc.SelectObject(ctx.brush0);
            dc.SelectObject(ctx.pen4);
        }
        
        // Draw completely matched column indicator (green dot) or empty column indicator (yellow dot)
        if (m_nMatrixDone && !m_bOnlyPcnt && ((xIndex - m_VisTopLeft.left) > 0))
        {
            if (m_pbGreenClms2[xIndex])
            {
                dc.SelectObject(ctx.pen5);
                dc.SelectObject(ctx.brush1);
                dc.Ellipse(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X, OFFSET_Y,
                           OFFSET_X + STEP_X + (xIndex - m_VisTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
            }
            if (m_engine.isEmptyCol2(xIndex))
            {
                dc.SelectObject(ctx.pen6);
                dc.SelectObject(ctx.brush2);
                dc.Ellipse(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X, OFFSET_Y,
                           OFFSET_X + STEP_X + (xIndex - m_VisTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
            }
        }
        
        // Draw the name of the column text
        dc.SetBkMode(TRANSPARENT);
        if (m_engine.isThisAKey(2, adjustedColIndex))
        {
            dc.SelectObject(ctx.font2B); // Bold if key
            dc.SetTextColor(RGB(0, 0, 170));
        }
        else
        {
            dc.SelectObject(ctx.font2);
            dc.SetTextColor(RGB(0, 0, 0));
        }
        dc.TextOutW(OFFSET_X + 5 + xIndex * STEP_X, -2 + OFFSET_Y + STEP_Y, m_Table2.Columns[adjustedColIndex]);
    }
}

void CChildView::paintMatrixCells(CDC& dc, PaintCtx& ctx)
{
    dc.SelectObject(ctx.pen2);
    if (m_nMatrixDone && !m_bOnlyPcnt)
    {
        // Set up DC for drawing percentages inside cells
        dc.SetBkMode(OPAQUE);
        dc.SelectObject(ctx.font3);
        int similarityPercentage; // Renamed from valSimil
        CString similarityStr; // Renamed from strSimil
        int adjustedRowIndex, adjustedColIndex; // Renamed from mx_y_adj, mx_x_adj
        
        if (m_nEffMax)
        {
            // Iterate over the visible cells of the matrix
            for (int yIndex = ctx.bnd_Y_min; yIndex <= ctx.bnd_Y_max - m_VisTopLeft.top; yIndex++)
            {
                for (int xIndex = ctx.bnd_X_min; xIndex <= ctx.bnd_X_max - m_VisTopLeft.left; xIndex++)
                {
                    dc.SelectObject(ctx.pen2);
                    adjustedRowIndex = yIndex + m_VisTopLeft.top;
                    adjustedColIndex = xIndex + m_VisTopLeft.left;
                    
                    // Calculate similarity score percentage and format as string
                    similarityPercentage = m_matrix.get(adjustedColIndex, adjustedRowIndex) * 100 / m_nEffMax;
                    similarityStr.Format(L"%i%%", similarityPercentage);
                    dc.SetBkMode(OPAQUE);
                    
                    // Determine background color based on matches and UI flags
                    if (!m_bSameNames || (m_Table1.Columns[adjustedRowIndex] == m_Table2.Columns[adjustedColIndex]))
                    {
                        if (similarityPercentage == 100)
                        {
                            if (m_engine.isEmptyCol1(adjustedRowIndex) || m_engine.isEmptyCol2(adjustedColIndex))
                                dc.SelectObject(ctx.brush2); // Yellow for empty but 100% match
                            else
                                dc.SelectObject(ctx.brush1); // Green for non-empty 100% match
                        }
                        if (similarityPercentage < 100)
                        {
                            if (similarityPercentage > m_nSldr) // Threshold slider check
                            {
                                dc.SelectObject(ctx.brush4); // Red/Orange for suspicious match
                            }
                            else
                            {
                                if (m_engine.isThisAKey(1, adjustedRowIndex) || m_engine.isThisAKey(2, adjustedColIndex))
                                    dc.SelectObject(ctx.brush6); // Grey for key intersections
                                else
                                    dc.SelectObject(ctx.brush0); // White for default
                            }
                        }
                    }
                    else
                    {
                        if (m_engine.isThisAKey(1, adjustedRowIndex) || m_engine.isThisAKey(2, adjustedColIndex))
                            dc.SelectObject(ctx.brush6); // Grey for key intersections
                        else
                            dc.SelectObject(ctx.brush0); // White for default
                    }
                    
                    // Highlight the clicked cell
                    if (adjustedRowIndex == m_CClickedCell.y && adjustedColIndex == m_CClickedCell.x)
                    {
                        dc.SelectObject(ctx.brush6);
                    }
                    
                    // Draw the cell's main rectangle
                    dc.Rectangle(OFFSET_X + xIndex * STEP_X, OFFSET_Y + yIndex * STEP_Y,
                                 1 + OFFSET_X + STEP_X + xIndex * STEP_X, 1 + OFFSET_Y + STEP_Y + yIndex * STEP_Y);
                    
                    // Draw X mark for actively marked/pinned cells
                    dc.SetBkMode(TRANSPARENT);
                    if (m_matrix.isMarked(adjustedColIndex, adjustedRowIndex))
                    {
                        dc.SelectObject(ctx.pen4);
                        dc.MoveTo(OFFSET_X + xIndex * STEP_X, OFFSET_Y + yIndex * STEP_Y);
                        dc.LineTo(OFFSET_X + STEP_X + xIndex * STEP_X, OFFSET_Y + STEP_Y + yIndex * STEP_Y);
                        dc.MoveTo(OFFSET_X + STEP_X + xIndex * STEP_X, OFFSET_Y + yIndex * STEP_Y);
                        dc.LineTo(OFFSET_X + xIndex * STEP_X, OFFSET_Y + STEP_Y + yIndex * STEP_Y);
                        dc.SelectObject(ctx.pen2);
                    }
                    
                    // Draw similarity mode specific highlights
                    if (m_bToDisplaySimilarClms && m_vecSimilaritiesAcrossTables[adjustedRowIndex].clm2 == adjustedColIndex)
                    {
                        dc.SetBkMode(TRANSPARENT);
                        dc.SelectObject(&m_KeyCurvePen);
                        dc.SelectObject(ctx.brush7);
                        dc.Rectangle(OFFSET_X + (xIndex)*STEP_X + 1, OFFSET_Y + (yIndex)*STEP_Y + 1,
                                     OFFSET_X + STEP_X + (xIndex)*STEP_X, OFFSET_Y + STEP_Y + (yIndex)*STEP_Y);
                    }
                    
                    // Output the percentage text
                    dc.SetTextColor(RGB(0, 0, 0));
                    dc.TextOutW(OFFSET_X + xIndex * STEP_X + 1, OFFSET_Y + yIndex * STEP_Y + 7, similarityStr);
                }
            }
            
            // Second pass for borders and highlights over cells
            dc.SetBkMode(TRANSPARENT);
            dc.SelectObject(GetStockObject(NULL_BRUSH));
            dc.SelectObject(ctx.pen3);
            for (int yIndex = ctx.bnd_Y_min; yIndex <= ctx.bnd_Y_max - m_VisTopLeft.top; yIndex++)
            {
                for (int xIndex = ctx.bnd_X_min; xIndex <= ctx.bnd_X_max - m_VisTopLeft.left; xIndex++)
                {
                    adjustedRowIndex = yIndex + m_VisTopLeft.top;
                    adjustedColIndex = xIndex + m_VisTopLeft.left;
                    
                    // Highlight intersection of columns with the same name (diagonal in 1-to-1 match)
                    if (m_Table1.Columns[adjustedRowIndex] == m_Table2.Columns[adjustedColIndex])
                    {
                        dc.Rectangle(OFFSET_X + xIndex * STEP_X, OFFSET_Y + yIndex * STEP_Y,
                                     1 + OFFSET_X + STEP_X + xIndex * STEP_X, 1 + OFFSET_Y + STEP_Y + yIndex * STEP_Y);
                    }
                    
                    // Draw hover/focus state for the previously clicked cell (to visualize movement)
                    if (adjustedRowIndex == m_nOldy && adjustedColIndex == m_nOldx)
                    {
                        dc.SelectObject(ctx.pen9);
                        dc.Rectangle(OFFSET_X + xIndex * STEP_X + 3, 1 + OFFSET_Y + STEP_Y + yIndex * STEP_Y - 4,
                                     1 + OFFSET_X + STEP_X + xIndex * STEP_X - 2,
                                     1 + OFFSET_Y + STEP_Y + yIndex * STEP_Y - 2);
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
    // Draw curved lines connecting similar columns across the two tables
    if (m_bToDisplaySimilarClms)
    {
        int xIndex, yIndex = 0; // Renamed from mx_x, mx_y to represent visual coordinates
        long maxHit = m_vecSimilaritiesAcrossTablesSorted[1].similarity;
        
        // Pass 1: Draw background similarity curves for all matched pairs
        for (int s_i = m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder; s_i >= 0; s_i--)
        {
            yIndex = m_vecSimilaritiesAcrossTablesSorted[s_i].clm1;
            xIndex = m_vecSimilaritiesAcrossTablesSorted[s_i].clm2;
            
            // Only draw if both coordinates represent valid columns and are within the visible bounds
            if ((yIndex * xIndex > 0) && ((yIndex - m_VisTopLeft.top) * (xIndex - m_VisTopLeft.left) > 0))
            {
                // Select a pen color based on the relative similarity score
                dc.SelectObject(&m_SimsPens[255 * m_vecSimilaritiesAcrossTablesSorted[s_i].similarity / maxHit]);
                
                // Define 4 control points for a Bezier curve connecting the two headers
                CPoint pt[4] = {
                    CPoint(OFFSET_X + STEP_X + 1, OFFSET_Y + (yIndex - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
                    CPoint(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X,
                           OFFSET_Y + (yIndex - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
                    CPoint(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X + STEP_X / 2,
                           OFFSET_Y + (yIndex - m_VisTopLeft.top) * STEP_Y),
                    CPoint(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X + STEP_X / 2, OFFSET_Y + STEP_Y)};
                dc.PolyBezier(pt, 4);
            }
        }
        
        // Pass 2: Draw primary key similarity curves (drawn on top of background curves)
        for (int s_i = m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder; s_i >= 0; s_i--)
        {
            yIndex = m_vecSimilaritiesAcrossTablesSorted[s_i].clm1;
            xIndex = m_vecSimilaritiesAcrossTablesSorted[s_i].clm2;
            if ((yIndex * xIndex > 0) && ((yIndex - m_VisTopLeft.top) * (xIndex - m_VisTopLeft.left) > 0))
            {
                // If both columns are marked as keys, draw them with the special key curve pen
                if (m_engine.isThisAKey(1, yIndex) && m_engine.isThisAKey(2, xIndex))
                {
                    dc.SelectObject(&m_SimsPens[255 * m_vecSimilaritiesAcrossTablesSorted[s_i].similarity / maxHit]);
                    CPoint pt[4] = {
                        CPoint(OFFSET_X + STEP_X + 1, OFFSET_Y + (yIndex - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
                        CPoint(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X,
                               OFFSET_Y + (yIndex - m_VisTopLeft.top) * STEP_Y + STEP_Y / 2),
                        CPoint(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X + STEP_X / 2,
                               OFFSET_Y + (yIndex - m_VisTopLeft.top) * STEP_Y),
                        CPoint(OFFSET_X + (xIndex - m_VisTopLeft.left) * STEP_X + STEP_X / 2, OFFSET_Y + STEP_Y)};
                    dc.SelectObject(&m_KeyCurvePen);
                    dc.PolyBezier(pt, 4);
                }
            }
        }
    }
}

void CChildView::pickFile(bool* pNewFile, ExcelConnector* pExcel, Table* pTable, CMFCRibbonComboBox* pSheetCombo,
                          CString* pFilename)
{
    *pNewFile = false;
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_TILL_IN_EXCEL));
        
    CString fileName;
    wchar_t* pFileBuffer = fileName.GetBuffer(FILE_LIST_BUFFER_SIZE); // Renamed from p to pFileBuffer
    
    // Display Open File Dialog
    CFileDialog dlgFile(TRUE);
    OPENFILENAME& ofn = dlgFile.GetOFN();
    ofn.lpstrFile = pFileBuffer;
    ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
    dlgFile.DoModal();
    fileName.ReleaseBuffer();
    
    // Extract individual filenames (handles potential multi-select buffer format)
    wchar_t* pBufEnd = pFileBuffer + FILE_LIST_BUFFER_SIZE - 2;
    wchar_t* startPointer = pFileBuffer; // Renamed from start to startPointer
    while ((pFileBuffer < pBufEnd) && (*pFileBuffer))
        pFileBuffer++;
        
    if (pFileBuffer > startPointer)
    {
        _tprintf(CMsg(IDS_PATH_TO_FILE), startPointer);
        pFileBuffer++;
        int fileCount = 1;
        while ((pFileBuffer < pBufEnd) && (*pFileBuffer))
        {
            startPointer = pFileBuffer;
            while ((pFileBuffer < pBufEnd) && (*pFileBuffer))
                pFileBuffer++;
            if (pFileBuffer > startPointer)
                _tprintf(_T("%2d. %s\r\n"), fileCount, startPointer);
            pFileBuffer++;
            fileCount++;
        }
    }
    
    // Close the currently opened Excel book in this connector
    pExcel->closeBook();
    
    // If a valid filename was selected, open it
    if (!(CString(fileName) == L""))
    {
        // Reset table column names
        for (int i = 0; i < 255; i++)
        {
            pTable->Columns[i] = "";
        }
        
        // Attempt to open the file via Excel COM application
        if (pExcel->openFile(fileName, m_App))
        {
            // Successfully opened, so populate the sheet selection combobox
            pSheetCombo->RemoveAllItems();
            CWorksheets& sheets = pExcel->getSheets();
            for (int i = 1; i <= sheets.get_Count(); i++)
            {
                if (CWorksheet tempSheet = sheets.get_Item(COleVariant((short)i)))
                {
                    pSheetCombo->AddItem(tempSheet.get_Name());
                }
                else
                {
                    break;
                }
            }
        }
    }
    
    // Store results
    *pFilename = fileName;
    *pNewFile = true;
    m_nUiToBeRefreshed = 3; // Trigger UI refresh for labels
    
    // Reset comparison matrix state
    if (m_nMatrixDone > 0)
    {
        m_matrix.clear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
        m_nMatrixDone = 0;
        m_OldCell.x = 0;
        m_OldCell.y = 0;
    }
    
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED));
        
    deleteAllKeys();
    this->Invalidate(); // Trigger a redraw of the view
}

void CChildView::OnPickFirstFile()
{
    bool* pNewFile = &m_bNewFile1;
    ExcelConnector* pExcel = &m_excel1;
    Table* pTable = &m_Table1;
    CMFCRibbonComboBox* pSheetCombo = m_pSheetCombo1;
    CString* pFilename = &m_szFilename1;

    pickFile(pNewFile, pExcel, pTable, pSheetCombo, pFilename);
}

void CChildView::OnPickSecondFile()
{
    bool* pNewFile = &m_bNewFile2;
    ExcelConnector* pExcel = &m_excel2;
    Table* pTable = &m_Table2;
    CMFCRibbonComboBox* pSheetCombo = m_pSheetCombo2;
    CString* pFilename = &m_szFilename2;

    pickFile(pNewFile, pExcel, pTable, pSheetCombo, pFilename);
}

void CChildView::OnCreateMatrix()
{
    if (m_bLockPrg1 || m_bLockPrg2)
    {
        MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
        return;
    }
    m_nMatrixDone = 0;
    m_nPrereqDone = 0;
    if (m_engine.areThereAnyKeys() == false)
    {
        MessageBox(CMsg(IDS_ATLEAST_ONE_KEY)); // CMsg(IDS_ATLEAST_ONE_KEY)
        return;
    }
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_COMPARISON_IN_PROGRESS)); // CMsg(IDS_COMPARISON_IN_PROGRESS)
    m_bWaitingForKeys = true;
    m_bKeys1done = false;
    m_bKeys2done = false;
    AfxBeginThread(CreateKeys1ThreadProc, this);
    AfxBeginThread(CreateKeys2ThreadProc, this);
    this->Invalidate();
    m_App.put_Visible(true);
    m_App.put_UserControl(TRUE);
}

void CChildView::OnUpdatePickFirstSheet(CCmdUI* pCmdUI)
{
    if (!(m_szFilename1 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSheetCombo1 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_PICK_FIRST_SHEET));
}

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
        if (m_nUiToBeRefreshed > 0)
            m_nUiToBeRefreshed -= 1;
    }
}

void CChildView::updateFileName(CCmdUI* pCmdUI, CString* pszFilename, int idString)
{
    if (m_nUiToBeRefreshed)
    {
        if (!(*pszFilename == ""))
        {
            CString s = *pszFilename;
            int origLen = s.GetLength();
            s = s.Right(20);
            s = (CString)CMsg(idString) + (origLen > 20 ? ".." : "") + s; // CMsg(idString)
            pCmdUI->SetText(s);
            pCmdUI->Enable(true);
            this->GetTopLevelFrame()->Invalidate();
        }
        else
        {
            pCmdUI->Enable(false);
        }
        if (m_nUiToBeRefreshed > 0)
            m_nUiToBeRefreshed -= 1;
    }
}

void CChildView::OnUpdateFilename1(CCmdUI* pCmdUI)
{
    updateFileName(pCmdUI, &m_szFilename1, IDS_1ST_FILE);
}

void CChildView::OnUpdateFilename2(CCmdUI* pCmdUI)
{
    updateFileName(pCmdUI, &m_szFilename2, IDS_2ND_FILE);
}

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
    if (nDelta != 0)
    {
        m_nVScrollPos += nDelta;
        m_VisTopLeft.top = m_nVScrollPos / STEP_Y;
        SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
        RECT rect;
        GetClientRect(&rect);
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
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(L"");
    return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}

void CChildView::OnUpdatePickSecondSheet(CCmdUI* pCmdUI)
{
    if (!(m_szFilename2 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSheetCombo2 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_PICK_SECOND_SHEET));
}

void CChildView::OnUpdateProgress1(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_PROGRESS2));
}

void CChildView::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
    int nDelta;
    switch (nSBCode)
    {
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
    if (nDelta != 0)
    {
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

void CChildView::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
    int nDelta;
    switch (nSBCode)
    {
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
    if (nDelta != 0)
    {
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

void CChildView::pickSheet(ExcelConnector* pExcel, Table* pTable, CMFCRibbonComboBox* pSheetCombo,
                           CMFCRibbonEdit* pSpinner_Names, CMFCRibbonEdit* pSpinner_Fdata, CMFCRibbonEdit* pRows,
                           CMFCRibbonEdit* pCols)
{
    int selectedSheetIndex = pSheetCombo->GetCurSel() + 1; // 1-based index (renamed from tmpWSN)
    CString selectedSheetName = pSheetCombo->GetEditText(); // Renamed from tmpWSS
    
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_PRELIM_CHK));
        
    if (selectedSheetIndex > 0)
    {
        pTable->WorkSheetNumber = selectedSheetIndex;
        long totalRows; // Renamed from iRows
        long totalCols; // Renamed from iCols
        
        // Select sheet in Excel via COM and retrieve actual dimensions
        pExcel->selectSheet(selectedSheetName, totalRows, totalCols);
        
        // Update table metadata
        pTable->MaxNumberOfRows = totalRows;
        pTable->MaxNumberOfCols = totalCols;
        pTable->NumberOfColumns = totalCols;
        pTable->NumberOfRows = totalRows;
        pTable->RowWithNames = 1;
        
        // Update ribbon UI controls with default row/column values
        CString strValue; // Renamed from tmps
        strValue.Format(_T("%d"), 1);
        pSpinner_Names->SetEditText(strValue); // Header row defaults to 1
        pTable->RowWithNames = 1;
        
        strValue.Format(_T("%d"), 2);
        pSpinner_Fdata->SetEditText(strValue); // First data row defaults to 2
        pTable->FirstRowWithData = 2;
        
        strValue.Format(_T("%d"), pTable->NumberOfRows);
        pRows->SetEditText(strValue);
        
        strValue.Format(_T("%d"), pTable->NumberOfColumns);
        pCols->SetEditText(strValue);
        
        // Setup initial cell dimensions and view extents for the scrollbar calculation
        m_nCellWidth = STEP_X;
        m_nCellHeight = STEP_Y;
        m_nRibbonWidth = 0;
        m_nViewWidth = STEP_X + OFFSET_X + ((pTable->NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
        m_nViewHeight = STEP_Y + OFFSET_Y + m_nCellHeight * (pTable->NumberOfColumns + 1);
        
        // Configure vertical scrollbar properties
        SCROLLINFO si;
        si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
        si.nMin = 0;
        si.nMax = m_nViewHeight - 1;
        si.nPos = m_nVScrollPos;
        si.nPage = m_nVPageSize;
        SetScrollInfo(SB_VERT, &si, TRUE);
        
        this->Invalidate(); // Redraw UI
        
        // Clear any existing comparison matrix as the active sheet has changed
        m_nMatrixDone = false;
        deleteAllKeys();
        if (m_nMatrixDone > 0)
        {
            m_matrix.clear(pTable->NumberOfColumns + 1, pTable->NumberOfRows + 1);
            m_nMatrixDone = 0;
            m_OldCell.x = 0;
            m_OldCell.y = 0;
            M_CCell.x = 0;
            M_CCell.y = 0;
        }
        
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED));
            
        m_engine.setTables(m_Table1, m_Table2);
        AfxBeginThread(makePrereq1ThreadProc, this); // Start parsing table headers in background
    }
}

void CChildView::OnPickFirstSheet()
{
    pickSheet(&m_excel1, &m_Table1, m_pSheetCombo1, m_pSpinner1_Names, m_pSpinner1_Fdata, m_pRows1, m_pCols1);
    updateCombos1();
}

void CChildView::OnSpin1Names()
{
    CString tmps = m_pSpinner1_Names->GetEditText();
    int tmpi = _ttoi(tmps);
    if (tmpi < 1)
        tmpi = 1;
    if (tmpi > 64)
        tmpi = 64;
    tmps.Format(_T("%d"), tmpi);
    m_pSpinner1_Names->SetEditText(tmps);
    m_Table1.RowWithNames = tmpi;
    updateCombos1();
    this->Invalidate();
}

void CChildView::OnUpdateSpin1Names(CCmdUI* pCmdUI)
{
    if (!(m_szFilename1 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSpinner1_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN1_NAMES));
}

void CChildView::OnUpdateSpin1Fdata(CCmdUI* pCmdUI)
{
    if (!(m_szFilename1 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSpinner1_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN1_FDATA));
}

void CChildView::OnSpin1Fdata()
{
    CString tmps = m_pSpinner1_Fdata->GetEditText();
    int tmpi = _ttoi(tmps);
    if (tmpi < 2)
        tmpi = 1;
    if (tmpi > 64)
        tmpi = 64;
    tmps.Format(_T("%d"), tmpi);
    m_pSpinner1_Fdata->SetEditText(tmps);
    m_Table1.FirstRowWithData = tmpi;
    m_engine.invalidatePrereq1();
}

void CChildView::updateCombos1()
{
    CString cellContent; // Renamed from szdata
    COleVariant vData;
    
    // Fetch header names from Excel and ensure they are unique to avoid map collisions
    for (int colIndex = 1; colIndex <= m_Table1.NumberOfColumns; colIndex++) // Renamed i to colIndex
    {
        cellContent = m_excel1.getCellValue(colIndex, m_Table1.RowWithNames);
        if (cellContent == "")
            cellContent = CMsg(IDS_NO_NAME); // Default name if header cell is empty
            
        // Check for duplicate names among previously read columns and append an index suffix if duplicate
        for (int prevCol = 1; prevCol < colIndex; prevCol++) // Renamed i1 to prevCol
        {
            if (cellContent == m_Table1.Columns[prevCol])
            {
                CString suffix;
                suffix.Format(L"[%i]", colIndex);
                cellContent += suffix;
                break;
            }
        }
        m_Table1.Columns[colIndex] = cellContent;
    }
}

void CChildView::OnPickSecondSheet()
{
    pickSheet(&m_excel2, &m_Table2, m_pSheetCombo2, m_pSpinner2_Names, m_pSpinner2_Fdata, m_pRows2, m_pCols2);
    updateCombos2();
}

void CChildView::updateCombos2()
{
    CString cellContent; // Renamed from szdata
    COleVariant vData;
    
    // Fetch header names from Excel and ensure they are unique to avoid map collisions
    for (int colIndex = 1; colIndex <= m_Table2.NumberOfColumns; colIndex++) // Renamed i to colIndex
    {
        cellContent = m_excel2.getCellValue(colIndex, m_Table2.RowWithNames);
        if (cellContent == "")
            cellContent = CMsg(IDS_NO_NAME); // Default name if header cell is empty
            
        // Check for duplicate names among previously read columns and append an index suffix if duplicate
        for (int prevCol = 1; prevCol < colIndex; prevCol++) // Renamed i1 to prevCol
        {
            if (cellContent == m_Table2.Columns[prevCol])
            {
                CString suffix;
                suffix.Format(L"[%i]", colIndex);
                cellContent += suffix;
                break;
            }
        }
        m_Table2.Columns[colIndex] = cellContent;
    }
}

void CChildView::OnUpdateSpin2Fdata(CCmdUI* pCmdUI)
{
    if (!(m_szFilename2 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSpinner2_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN2_FDATA));
}

void CChildView::OnSpin2Fdata()
{
    CString tmps = m_pSpinner2_Fdata->GetEditText();
    int tmpi = _ttoi(tmps);
    if (tmpi < 2)
        tmpi = 1;
    if (tmpi > 64)
        tmpi = 64;
    tmps.Format(_T("%d"), tmpi);
    m_pSpinner2_Fdata->SetEditText(tmps);
    m_Table2.FirstRowWithData = tmpi;
    m_engine.invalidatePrereq2();
}

void CChildView::OnUpdateSpin2Names(CCmdUI* pCmdUI)
{
    if (!(m_szFilename2 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSpinner2_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN2_NAMES));
}

void CChildView::OnSpin2Names()
{
    CString tmps = m_pSpinner2_Names->GetEditText();
    int tmpi = _ttoi(tmps);
    if (tmpi < 1)
        tmpi = 1;
    if (tmpi > 64)
        tmpi = 64;
    tmps.Format(_T("%d"), tmpi);
    m_pSpinner2_Names->SetEditText(tmps);
    m_Table2.RowWithNames = tmpi;
    updateCombos2();
    this->Invalidate();
}

void CChildView::OnLButtonDblClk(UINT nFlags, CPoint point)
{
    if (m_bLockPrg1 || m_bLockPrg2)
    {
        MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
        return;
    }
    if (m_nMatrixDone && (M_CCell.y <= m_Table1.NumberOfColumns && M_CCell.x <= m_Table2.NumberOfColumns))
    {
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
        m_bLockPrg2 = true;
        int mx_X_max = m_Table2.NumberOfColumns;
        m_matrix.setMarked(M_CCell.x, M_CCell.y);
        this->Invalidate();
        AfxBeginThread(MyThreadProc3, this);
    }
    CWnd::OnLButtonDblClk(nFlags, point);
}

bool CChildView::checkKeysUniqueness1()
{
    m_bLockPrg1 = true;
    bool result = m_engine.checkKeysUniqueness1();
    m_bLockPrg1 = false;
    return result;
}

bool CChildView::checkKeysUniqueness2()
{
    m_bLockPrg2 = true;
    bool result = m_engine.checkKeysUniqueness2();
    m_bLockPrg2 = false;
    return result;
}

void CChildView::firstPass()
{
    m_bLockPrg1 = true;
    m_engine.setTables(m_Table1, m_Table2);
    m_engine.firstPass(m_matrix, m_bAutoMark, m_bIn2file, m_pbGreenClms1, m_pbGreenClms2, m_nEffMax, m_bDoAutoMark);
    m_nMatrixDone++;
    m_bLockPrg1 = false;
}

int CChildView::createKeyArrays1()
{
    m_engine.setTables(m_Table1, m_Table2);
    m_engine.m_bUseIndexes = m_bUseIndexes;
    m_bLockPrg1 = true;
    int result = m_engine.createKeyArrays1();
    m_bLockPrg1 = false;
    return result;
}

int CChildView::createKeyArrays2()
{
    m_engine.setTables(m_Table1, m_Table2);
    m_engine.m_bUseIndexes = m_bUseIndexes;
    m_bLockPrg2 = true;
    int result = m_engine.createKeyArrays2();
    m_bLockPrg2 = false;
    return result;
}

void CChildView::OnMouseMove(UINT nFlags, CPoint point)
{
    m_OldCell.x = M_CCell.x;
    m_OldCell.y = M_CCell.y;
    
    // Calculate current hovered column index (Table 2) based on mouse X coordinate
    if (point.x > OFFSET_X + STEP_X)
    {
        M_CCell.x = (point.x - OFFSET_X) / STEP_X + m_VisTopLeft.left;
    }
    else
    {
        M_CCell.x = 0;
    }
    
    // Calculate current hovered row index (Table 1) based on mouse Y coordinate
    if (point.y > OFFSET_Y + STEP_Y)
    {
        M_CCell.y = (point.y - OFFSET_Y) / STEP_Y + m_VisTopLeft.top;
    }
    else
    {
        M_CCell.y = 0;
    }
    
    // If in similarity display mode, snap the X coordinate to the most similar column
    if (m_bToDisplaySimilarClms)
    {
        if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns)
        {
            int similarColIndex = m_vecSimilaritiesAcrossTables[M_CCell.y].clm2; // Renamed from tmpCellx
            if (similarColIndex > 0 && similarColIndex <= m_Table2.NumberOfColumns)
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
    
    // Update the status bar with the current hovered coordinates
    if (M_CCell.x * M_CCell.y > 0)
    {
        CString statusText; // Renamed from s, sx and sy removed as they were redundant
        statusText.Format(CMsg(IDS_COORDS), M_CCell.y, M_CCell.x);
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(statusText);
    }
    
    // Only trigger redraw logic if the hovered cell has changed
    if (!(m_OldCell.x == M_CCell.x) || !(m_OldCell.y == M_CCell.y))
    {
        if (!m_bForceNotOnlyPcnt)
        {
            m_bOnlyPcnt = true; // Optimization flag to avoid full redraw
        }
        else
        {
            m_bOnlyPcnt = false;
            m_bForceNotOnlyPcnt = false;
        }
        
        RECT invalidRect; // Renamed from rct
        
        // Invalidate top-left corner area
        invalidRect.left = 0;
        invalidRect.top = 0;
        invalidRect.right = OFFSET_X + STEP_X / 2;
        invalidRect.bottom = OFFSET_Y + STEP_Y / 2;
        this->InvalidateRect(&invalidRect, 1);
        
        // Invalidate current hovered cell headers to show focus highlights
        if (M_CCell.y > 0 && M_CCell.y <= m_Table1.NumberOfColumns && M_CCell.x > 0 &&
            M_CCell.x <= m_Table2.NumberOfColumns)
        {
            // Invalidate column header area
            invalidRect.left = OFFSET_X + (M_CCell.x - m_VisTopLeft.left) * STEP_X + 1;
            invalidRect.top = 2;
            invalidRect.right = 1 + OFFSET_X + STEP_X + (M_CCell.x - m_VisTopLeft.left) * STEP_X;
            invalidRect.bottom = OFFSET_Y + STEP_Y;
            this->InvalidateRect(&invalidRect, 0);
            
            // Invalidate row header area
            invalidRect.left = 2;
            invalidRect.top = OFFSET_Y + (M_CCell.y - m_VisTopLeft.top) * STEP_Y + 1;
            invalidRect.right = OFFSET_X + STEP_X;
            invalidRect.bottom = 1 + OFFSET_Y + (M_CCell.y - m_VisTopLeft.top) * STEP_Y + STEP_Y;
            this->InvalidateRect(&invalidRect, 0);
        }
        
        // Invalidate previously hovered cell headers to clear their highlights
        if (m_OldCell.y > 0 && m_OldCell.y <= m_Table1.NumberOfColumns && m_OldCell.x > 0 &&
            m_OldCell.x <= m_Table2.NumberOfColumns)
        {
            // Clear column header highlight
            invalidRect.left = OFFSET_X + (m_OldCell.x - m_VisTopLeft.left) * STEP_X + 1;
            invalidRect.top = 2;
            invalidRect.right = 1 + OFFSET_X + STEP_X + (m_OldCell.x - m_VisTopLeft.left) * STEP_X;
            invalidRect.bottom = OFFSET_Y + STEP_Y;
            this->InvalidateRect(&invalidRect, 1);
            
            // Clear row header highlight
            invalidRect.left = 2;
            invalidRect.top = OFFSET_Y + (m_OldCell.y - m_VisTopLeft.top) * STEP_Y + 1;
            invalidRect.right = OFFSET_X + STEP_X;
            invalidRect.bottom = 1 + OFFSET_Y + (m_OldCell.y - m_VisTopLeft.top) * STEP_Y + STEP_Y;
            this->InvalidateRect(&invalidRect, 1);
        }
        
        m_bOnlyPcnt = false;
        m_bForceNotOnlyPcnt = true;
        
        // Reset coordinates if we hovered outside the valid bounds
        if (M_CCell.x * M_CCell.y == 0)
        {
            M_CCell.x = 0;
            M_CCell.y = 0;
        }
    }
    this->SetFocus();
}

void CChildView::OnSimilarityThreshold()
{
    m_nSldr = m_pSlider->GetPos();
    this->Invalidate();
    CString s;
    CString sx;
    s = m_szRsltTxt;
    sx.Format(CMsg(IDS_MARK_SUSP_INTERS), m_pSlider->GetPos()); // CMsg(IDS_MARK_SUSP_INTERS)
    s = sx + L" %";
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(s);
}

void CChildView::OnUpdateSimilarityThreshold(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSlider = DYNAMIC_DOWNCAST(CMFCRibbonSlider, m_pRibbon->FindByID(ID_SIMILARITY_THRESHOLD));
    if (m_pSlider->GetPos() == 0)
        m_pSlider->SetPos(m_nSldr);
}

void CChildView::OnMarkInFile1()
{
    m_bIn1file = !m_bIn1file;
}

void CChildView::OnUpdateMarkInFile1(CCmdUI* pCmdUI)
{
    if (!(m_szFilename1 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    pCmdUI->SetCheck(m_bIn1file);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pMarkIn1 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_MARK_IN_FILE1));
}

void CChildView::OnMarkInFile2()
{
    m_bIn2file = !m_bIn2file;
}

void CChildView::OnUpdateMarkInFile2(CCmdUI* pCmdUI)
{
    if (!(m_szFilename2 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    pCmdUI->SetCheck(m_bIn2file);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pMarkIn2 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_MARK_IN_FILE2));
}

void CChildView::OnSuggestKeys()
{
    if (m_bLockPrg1 || m_bLockPrg2)
    {
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

void CChildView::markIn1(int row, int clm)
{
    CString cnv = convertR1C1(row, clm);
    m_excel1.markCellRange(
        cnv, cnv,
        RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
}

void CChildView::markIn2(int row, int clm)
{
    CString cnv = convertR1C1(row, clm);
    m_excel2.markCellRange(
        cnv, cnv,
        RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
}

void CChildView::initScrollBars()
{
    SCROLLINFO ScrollInfo;
    ScrollInfo.cbSize = sizeof(ScrollInfo); // size of this structure
    ScrollInfo.fMask = SIF_ALL;             // parameters to set
    ScrollInfo.nMin = 0;                    // minimum scrolling position
    ScrollInfo.nMax = 100;                  // maximum scrolling position
    ScrollInfo.nPage = 20;                  // the page size of the scroll box
    ScrollInfo.nPos = 50;
    // initial position of the scroll box
    //ScrollInfo.nTrackPos = 0;                   // immediate position of a scroll box
    this->SetScrollInfo(SB_HORZ, &ScrollInfo);
}

void CChildView::OnSize(UINT nType, int cx, int cy)
{
    CWnd::OnSize(nType, cx, cy);
    //
    // Set the horizontal scrolling parameters.
    //
    int nHScrollMax = 0;
    m_nHScrollPos = m_nHPageSize = 0;
    if (cx < m_nViewWidth)
    {
        nHScrollMax = m_nViewWidth - 1;
        m_nHPageSize = cx;
        m_nHScrollPos = min(m_nHScrollPos, m_nViewWidth - m_nHPageSize - 1);
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
    if (cy < m_nViewHeight)
    {
        nVScrollMax = m_nViewHeight - 1;
        m_nVPageSize = cy;
        m_nVScrollPos = min(m_nVScrollPos, m_nViewHeight - m_nVPageSize - 1);
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

void CChildView::OnUpdateProgress2(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pProgressBar2 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_PROGRESS3));
    // Emergency update of the container for found differences
    m_pFoundDifferences = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_DIFFS_LIST));
    m_pToFront = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_PUT_TO_FRONT));
}

void CChildView::OnUpdateVerifyKeys(CCmdUI* pCmdUI)
{
    pCmdUI->SetCheck(m_bVerifyKeys);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pVerifyKeys = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_VERIFY_KEYS));
}

void CChildView::OnVerifyKeys()
{
    m_bVerifyKeys = !m_bVerifyKeys;
}

void CChildView::OnUpdateSuggestKeys(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(true);
    //pCmdUI->SetText(m_bUseIndexes ? L"Sestavit kl��" : L"Naj�t kl��");
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pButton2 = DYNAMIC_DOWNCAST(CMFCRibbonButton, m_pRibbon->FindByID(ID_SUGGEST_KEYS));
}

void CChildView::OnSameNamesOnly()
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

void CChildView::OnUpdateSameNamesOnly(CCmdUI* pCmdUI)
{
    pCmdUI->SetCheck(m_bSameNames);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pSameNames = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_SAME_NAMES_ONLY));
}

UINT MyThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->firstPass();
    AfxEndThread(0);
    return 0;
}

afx_msg LRESULT CChildView::OnCmUpdateProgress(WPARAM wParam, LPARAM lParam)
{
    if ((UINT)lParam > 99)
    {
        m_pProgressBar1->SetPos(0);
        m_bLockPrg1 = false;
        this->Invalidate();
        if (m_bDoAutoMark)
        {
            if (g_pMainFrame)
                g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
            resolveAutoMark();
            if (g_pMainFrame)
                g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
        }
    }
    else
    {
        m_pProgressBar1->SetPos((UINT)lParam);
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmUpdateProgress2(WPARAM wParam, LPARAM lParam)
{
    if ((UINT)lParam > 99)
    {
        m_pProgressBar2->SetPos(0);
        m_bLockPrg2 = false;
    }
    else
    {
        m_pProgressBar2->SetPos((UINT)lParam);
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmUpdateProgress3(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar1->SetPos((UINT)lParam);
    return 0;
}

afx_msg LRESULT CChildView::OnCmMarkingReady(WPARAM wParam, LPARAM lParam)
{
    HWND hWnd = this->GetSafeHwnd();
    long totalRowsTable1; // Renamed from nor
    int currentProgress = 0, lastReportedProgress = 0; // Renamed from prgHlpr, prgHlpr0
    CString diffEntry, diffPart1, diffPart2, keyPart; // Renamed from fndDfrnc, fndDfrnc1, fndDfrnc2, selKey
    int differencesCount = 0; // Renamed from dfrnCntr
    long diffRowInTable2; // Renamed from dfrncRow2
    CString selectionStartCell = L""; // Renamed from starts
    CString selectionEndCell = L"";   // Renamed from ends
    
    BeginWaitCursor();
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING));
        
    m_pProgressBar1->SetPos(0);
    m_pFoundDifferences->RemoveAllItems();
    m_pFoundDifferences->SetEditText(L"");
    
    // Process Table 1 to identify and display differences, and optionally highlight them in Excel
    totalRowsTable1 = m_Table1.NumberOfRows + 1;
    for (int rowTable1 = 1; rowTable1 < totalRowsTable1; rowTable1++) // Renamed i1 to rowTable1
    {
        // Calculate and report progress
        currentProgress = (rowTable1 * 100) / totalRowsTable1;
        if (currentProgress > lastReportedProgress)
        {
            SendMessage(CM_UPDATE_PROGRESS2, 0, currentProgress);
            lastReportedProgress = currentProgress;
        }
        
        diffRowInTable2 = m_pnFoundDifferences[rowTable1];
        if (diffRowInTable2 > 0)
        {
            // We found a difference between the two tables
            if (++differencesCount < 500) // Limit the number of differences displayed to avoid overwhelming UI
            {
                // Format the difference string for Table 1
                diffPart1 = L"";
                diffPart1.Format(L"(1r%i):", rowTable1);
                diffPart1 += m_excel1.getCellValue(m_nOldy, rowTable1);
                diffPart1 = diffPart1.Left(26);
                
                // Format the difference string for Table 2
                diffPart2 = L"";
                diffPart2.Format(L"   (2r%i):", diffRowInTable2);
                diffPart2 += m_excel2.getCellValue(m_nOldx, diffRowInTable2);
                diffPart2 = diffPart2.Left(26);
                
                // Combine with the primary key string
                keyPart = L"";
                keyPart.Format(L"%s%s   (key): %s", diffPart1.GetString(), diffPart2.GetString(), m_engine.getKeyStr1(rowTable1).GetString());
                diffEntry = keyPart.Left(54);
                
                m_pFoundDifferences->AddItem((LPCTSTR)diffEntry);
            }
        }
        
        // Handle marking background colors in Excel for Table 1
        if (m_bIn1file)
        {
            if (m_pbMarkIn1Arr[rowTable1])
            {
                // Extend the continuous selection range
                if (selectionStartCell == L"")
                {
                    selectionStartCell = convertR1C1(rowTable1, m_nOldy);
                }
                selectionEndCell = convertR1C1(rowTable1, m_nOldy);
            }
            else
            {
                // The continuous block ended, apply the color to the range
                if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
                {
                    m_excel1.markCellRange(selectionStartCell, selectionEndCell,
                                           RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green,
                                               m_Palette[m_nChosenColor1].blue));
                    selectionStartCell = L"";
                    selectionEndCell = L"";
                }
            }
        }
    }
    
    // Process any remaining marking ranges for Table 1
    if (m_bIn1file && !(selectionStartCell == L"") && !(selectionEndCell == L""))
    {
        m_excel1.markCellRange(
            selectionStartCell, selectionEndCell,
            RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
        selectionStartCell = L"";
        selectionEndCell = L"";
    }
    
    // Handle marking background colors in Excel for Table 2
    if (m_bIn2file)
    {
        long totalRowsTable2 = m_Table2.NumberOfRows + 1; // Renamed from nor
        for (int rowTable2 = 1; rowTable2 < totalRowsTable2; rowTable2++) // Renamed i2 to rowTable2
        {
            // Calculate and report progress
            currentProgress = (rowTable2 * 100) / totalRowsTable2;
            if (currentProgress > lastReportedProgress)
            {
                SendMessage(CM_UPDATE_PROGRESS2, 0, currentProgress);
                lastReportedProgress = currentProgress;
            }
            
            if (m_pbMarkIn2Arr[rowTable2])
            {
                // Extend the continuous selection range
                if (selectionStartCell == L"")
                {
                    selectionStartCell = convertR1C1(rowTable2, m_nOldx);
                }
                selectionEndCell = convertR1C1(rowTable2, m_nOldx);
            }
            else
            {
                // The continuous block ended, apply the color to the range
                if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
                {
                    m_excel2.markCellRange(selectionStartCell, selectionEndCell,
                                           RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green,
                                               m_Palette[m_nChosenColor2].blue));
                    selectionStartCell = L"";
                    selectionEndCell = L"";
                }
            }
        }
        
        // Process any remaining marking ranges for Table 2
        if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
        {
            m_excel2.markCellRange(
                selectionStartCell, selectionEndCell,
                RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
            selectionStartCell = L"";
            selectionEndCell = L"";
        }
    }
    
    SendMessage(CM_UPDATE_PROGRESS2, 0, 100);
    m_bLockPrg2 = false;
    
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE));
        
    EndWaitCursor();
    DrainMsgQueue();
    return 0;
}

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
        pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE)); // CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE)
        return 0;
    }
    AfxEndThread(0);
    return 0;
}

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
        s.Format(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow,
                 nu->secondRow); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS)
        pWnd->MessageBox(s);
        pWnd->m_bLockPrg1 = false;
        return 0;
    }
    pWnd->m_bUniqueKeys1 = true;
    AfxEndThread(0);
    return 0;
}

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
        s.Format(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow,
                 nu->secondRow); // CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS)
        pWnd->MessageBox(s);
        pWnd->m_bLockPrg2 = false;
        return 0;
    }
    pWnd->m_bUniqueKeys2 = true;
    AfxEndThread(0);
    return 0;
}

UINT makePrereq1ThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->m_engine.makePrereq1();
    AfxEndThread(0);
    return 0;
}

UINT makePrereq2ThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->m_engine.makePrereq2();
    AfxEndThread(0);
    return 0;
}

UINT MyThreadProc3(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->markInFiles();
    AfxEndThread(0);
    return 0;
}

void CChildView::markInFiles()
{
    m_bLockPrg2 = true;
    int currentProgress = 0, lastReportedProgress = 0; // Renamed from prgHlpr, prgHlpr0
    int focusedColIndex = M_CCell.x; // Renamed from cx
    int focusedRowIndex = M_CCell.y; // Renamed from cy
    
    m_nOldy = focusedRowIndex;
    m_nOldx = focusedColIndex;
    
    // Initialize marking arrays to track which rows need highlighting in Excel
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
    
    // Find differences between the selected columns in Table 1 and Table 2 using primary keys
    int rowsProcessed = m_Table1.FirstRowWithData - 1; // Renamed from i1
    for (const auto& [key, keyRow1] : m_engine.getMap1())
    {
        rowsProcessed++;
        
        // Report progress back to the main UI thread
        lastReportedProgress = 100 * rowsProcessed / m_Table1.NumberOfRows;
        if (lastReportedProgress > currentProgress)
        {
            currentProgress = lastReportedProgress;
            PostMessage(CM_UPDATE_PROGRESS3, 0, currentProgress);
        }
        
        // Check if the key exists in Table 2
        if (auto it2 = m_engine.getMap2().find(key); it2 != m_engine.getMap2().end())
        {
            const long keyRow2 = it2->second;
            
            // Compare the cell values directly across the selected columns
            if (!(m_excel1.getCellValue(focusedRowIndex, keyRow1) == m_excel2.getCellValue(focusedColIndex, keyRow2)))
            {
                // Record the difference
                m_pnFoundDifferences[keyRow1] = keyRow2;
                
                // Mark for highlighting if enabled by the user
                if (m_bIn1file)
                    m_pbMarkIn1Arr[keyRow1] = true;
                if (m_bIn2file)
                    m_pbMarkIn2Arr[keyRow2] = true;
            }
        }
    }
    
    // Signal the UI thread that marking calculation is ready
    PostMessage(CM_MARKING_READY, 0, 0);
    m_bLockPrg2 = false;
}

void CChildView::OnColorPicker1()
{
    COLORREF i = (int)m_pColorPicker1->GetSelectedItem();
    m_nChosenColor1 = i;
}

void CChildView::OnUpdateColorPicker1(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(true);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pColorPicker1 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_pRibbon->FindByID(ID_COLOR_PICKER1));
}

void CChildView::OnColorPicker2()
{
    COLORREF i = (int)m_pColorPicker2->GetSelectedItem();
    m_nChosenColor2 = i;
}

void CChildView::OnUpdateColorPicker2(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(true);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pColorPicker2 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_pRibbon->FindByID(ID_COLOR_PICKER2));
}

void CChildView::OnAutoMark()
{
    m_bAutoMark = !m_bAutoMark;
}

void CChildView::OnUpdateAutoMark(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(true);
    pCmdUI->SetCheck(m_bAutoMark);
}

void CChildView::resolveAutoMark()
{
    m_bDoAutoMark = false;
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_DURING_MARKING_THREAD_BLOCKED));
        
    m_bLockPrg2 = true;
    HWND hWnd = this->GetSafeHwnd();
    int progressCol = 0, lastProgressCol = 0, progressRow = 0, lastProgressRow = 0; // Renamed from prgHlpr_x, etc.
    CString selectionStartCell = L""; // Renamed from starts
    CString selectionEndCell = L"";   // Renamed from ends
    
    m_pProgressBar1->SetPos(0);
    BeginWaitCursor();
    
    // Iterate through all columns to find identically named columns across tables
    for (int colIndex1 = 1; colIndex1 <= m_Table1.NumberOfColumns; colIndex1++) // Renamed c1 to colIndex1
    {
        lastProgressCol = 90 * colIndex1 / m_Table1.NumberOfColumns;
        if (lastProgressCol > progressCol)
        {
            progressCol = lastProgressCol;
            PostMessage(CM_UPDATE_PROGRESS, 0, progressCol);
        }
        
        for (int colIndex2 = 1; colIndex2 <= m_Table2.NumberOfColumns; colIndex2++) // Renamed c2 to colIndex2
        {
            // If the column names match, proceed to auto-mark their differences
            if (m_Table1.Columns[colIndex1] == m_Table2.Columns[colIndex2])
            {
                // Handle Auto-Marking in Table 1
                if (m_bIn1file)
                {
                    progressRow = 0;
                    lastProgressRow = 0;
                    for (long row1 = m_Table1.FirstRowWithData; row1 <= m_Table1.NumberOfRows; row1++) // Renamed r1 to row1
                    {
                        lastProgressRow = 100 * row1 / m_Table1.NumberOfRows;
                        if (lastProgressRow > progressRow + 10)
                        {
                            progressRow = lastProgressRow;
                            PostMessage(CM_UPDATE_PROGRESS2, 0, progressRow);
                        }
                        
                        // Check if the cell has different content across the mapped key
                        if (m_engine.getMainChar1(row1, colIndex1) == 1)
                        {
                            if (selectionStartCell == L"")
                            {
                                selectionStartCell = convertR1C1(row1, colIndex1);
                            }
                            selectionEndCell = convertR1C1(row1, colIndex1);
                        }
                        else
                        {
                            if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
                            {
                                m_excel1.markCellRange(selectionStartCell, selectionEndCell,
                                                       RGB(m_Palette[m_nChosenColor1].red,
                                                           m_Palette[m_nChosenColor1].green,
                                                           m_Palette[m_nChosenColor1].blue));
                                selectionStartCell = L"";
                                selectionEndCell = L"";
                            }
                        }
                    }
                    if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
                    {
                        m_excel1.markCellRange(selectionStartCell, selectionEndCell,
                                               RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green,
                                                   m_Palette[m_nChosenColor1].blue));
                        selectionStartCell = L"";
                        selectionEndCell = L"";
                    }
                }
                
                selectionStartCell = L"";
                selectionEndCell = L"";
                
                // Handle Auto-Marking in Table 2
                if (m_bIn2file)
                {
                    progressRow = 0;
                    lastProgressRow = 0;
                    for (long row2 = m_Table2.FirstRowWithData; row2 <= m_Table2.NumberOfRows; row2++) // Renamed r2 to row2
                    {
                        lastProgressRow = 100 * row2 / m_Table2.NumberOfRows;
                        if (lastProgressRow > progressRow + 10)
                        {
                            progressRow = lastProgressRow;
                            PostMessage(CM_UPDATE_PROGRESS2, 0, progressRow);
                        }
                        
                        // Check if the cell has different content across the mapped key
                        if (m_engine.getMainChar2(row2, colIndex2) == 1)
                        {
                            if (selectionStartCell == L"")
                            {
                                selectionStartCell = convertR1C1(row2, colIndex2);
                            }
                            selectionEndCell = convertR1C1(row2, colIndex2);
                        }
                        else
                        {
                            if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
                            {
                                m_excel2.markCellRange(selectionStartCell, selectionEndCell,
                                                       RGB(m_Palette[m_nChosenColor2].red,
                                                           m_Palette[m_nChosenColor2].green,
                                                           m_Palette[m_nChosenColor2].blue));
                                selectionStartCell = L"";
                                selectionEndCell = L"";
                            }
                        }
                    }
                    if (!(selectionStartCell == L"") && !(selectionEndCell == L""))
                    {
                        m_excel2.markCellRange(selectionStartCell, selectionEndCell,
                                               RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green,
                                                   m_Palette[m_nChosenColor2].blue));
                        selectionStartCell = L"";
                        selectionEndCell = L"";
                    }
                }
            }
        }
    }
    
    // Highlight missing keys entirely in Table 1
    for (long row1 = m_Table1.FirstRowWithData; row1 <= m_Table1.NumberOfRows; row1++)
    {
        if (m_engine.isKeyMissing1(row1))
        {
            selectionStartCell = convertR1C1(row1, 1);
            selectionEndCell = convertR1C1(row1, m_Table1.NumberOfColumns);
            m_excel1.markCellRange(
                selectionStartCell, selectionEndCell,
                RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue));
        }
    }
    
    // Highlight missing keys entirely in Table 2
    for (long row2 = m_Table2.FirstRowWithData; row2 <= m_Table2.NumberOfRows; row2++)
    {
        if (m_engine.isKeyMissing2(row2))
        {
            selectionStartCell = convertR1C1(row2, 1);
            selectionEndCell = convertR1C1(row2, m_Table2.NumberOfColumns);
            m_excel2.markCellRange(
                selectionStartCell, selectionEndCell,
                RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue));
        }
    }
    
    PostMessage(CM_UPDATE_PROGRESS, 0, 100);
    PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
    m_bLockPrg2 = false;
    EndWaitCursor();
    
    if (g_pMainFrame)
        g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE));
        
    DrainMsgQueue();
}

void CChildView::DrainMsgQueue(void)
{
    MSG msg = {0};
    HWND hWnd = this->GetSafeHwnd();
    while (PeekMessage(&msg, hWnd, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE))
        ;
}

void CChildView::OnDiffslist()
{
    // there is no required answer for this event - at least for now
}

void CChildView::OnUpdateDiffslist(CCmdUI* pCmdUI)
{
    // there is no required answer for this event - at least for now
}

void CChildView::OnGotoDiffInFile1()
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

void CChildView::OnGotoDiffInFile2()
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

void CChildView::OnBringExcelToFront()
{
    m_bToFront = !m_bToFront;
}

void CChildView::OnUpdateBringExcelToFront(CCmdUI* pCmdUI)
{
    pCmdUI->SetCheck(m_bToFront);
}

void CChildView::suggestKeys1()
{
    m_bLockPrg1 = true;
    m_keyFinder.suggestKeys1();
}

void CChildView::suggestKeys2()
{
    m_bLockPrg2 = true;
    m_keyFinder.suggestKeys2();
}

void CChildView::clearPossibleKeys()
{
    m_keyFinder.clearPossibleKeys();
}

UINT SuggestKeys1ThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->suggestKeys1();
    AfxEndThread(0);
    return 0;
}

UINT SuggestKeys2ThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->suggestKeys2();
    AfxEndThread(0);
    return 0;
}

UINT MutualCheckThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->mutualCheck();
    AfxEndThread(0);
    return 0;
}

UINT FindSimsThreadProc(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->findSims();
    AfxEndThread(0);
    return 0;
}

UINT FindSimsThreadProc1(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->findSims1();
    AfxEndThread(0);
    return 0;
}

UINT FindSimsThreadProc2(LPVOID pParam)
{
    CChildView* pWnd = static_cast<CChildView*>(pParam);
    pWnd->findSims2();
    AfxEndThread(0);
    return 0;
}

bool CChildView::mutualCheck()
{
    m_keyFinder.resetBestKeyComb();
    int mutualCheckResult = 0; // Renamed from tmpRslt
    
    // Reset key progress bars before starting the cascade check
    if (m_pKeyProgressBar1 && m_pKeyProgressBar2)
    {
        PostMessage(CM_UPDATE_KEYPROGRESS1, 0, 0);
        PostMessage(CM_UPDATE_KEYPROGRESS2, 0, 0);
    }
    
    // Basic validation to ensure at least one table has possible keys
    if (m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, 0) == 0 &&
        m_keyFinder.getNumberOfPossibleKeys(2, SUGKEYS, 0) == 0)
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
    
    int currentKeyIndex = 0; // Renamed from m_i
    m_bLockPrg1 = true;
    int currentProgress = 0, lastReportedProgress = 0; // Renamed from prgHlpr, prgHlpr0
    int similarityOrder = 1; // Renamed from order
    int possibleKeyCount1 = m_keyFinder.getPossibleKeyCounter1(); // Renamed from pkCnt1
    
    // First pass: Verify the top candidate keys
    while (currentKeyIndex <= possibleKeyCount1 && mutualCheckResult < 100)
    {
        lastReportedProgress = 100 * currentKeyIndex / (possibleKeyCount1 > 0 ? possibleKeyCount1 : 1);
        if (lastReportedProgress > currentProgress)
        {
            currentProgress = lastReportedProgress;
            PostMessage(CM_UPDATE_PROGRESS, 0, currentProgress);
            PostMessage(CM_UPDATE_KEYPROGRESS1, 0, currentProgress);
        }
        
        if (m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, currentKeyIndex) == similarityOrder)
        {
            mutualCheckResult = m_keyFinder.checkKeys(currentKeyIndex);
        }
        else
        {
            break;
        }
        currentKeyIndex++;
    }
    
    similarityOrder++;
    int maxSimilarityOrder = m_keyFinder.getNumberOfPossibleKeys(1, SUGKEYS, (possibleKeyCount1 - 1 >= 0 ? possibleKeyCount1 - 1 : 0)); // Renamed from maxOrder
    
    // Second pass: Exhaustively verify remaining keys of lower similarity orders if first pass was not confident enough
    while (mutualCheckResult < 90 && similarityOrder <= maxSimilarityOrder)
    {
        while (currentKeyIndex <= possibleKeyCount1 && mutualCheckResult < 90)
        {
            lastReportedProgress = 100 * currentKeyIndex / (possibleKeyCount1 > 0 ? possibleKeyCount1 : 1);
            if (lastReportedProgress > currentProgress)
            {
                currentProgress = lastReportedProgress;
                PostMessage(CM_UPDATE_PROGRESS, 0, currentProgress);
                PostMessage(CM_UPDATE_KEYPROGRESS1, 0, currentProgress);
            }
            
            if (m_keyFinder.getNumberOfPossibleKeys(1, similarityOrder, currentKeyIndex) == similarityOrder)
            {
                mutualCheckResult = m_keyFinder.checkKeys(currentKeyIndex);
            }
            else
            {
                break;
            }
            currentKeyIndex++;
        }
        similarityOrder++;
    }
    
    if (mutualCheckResult)
    {
        PostMessage(CM_KEYS_FOUND, 0, 0);
        return true;
    }
    
    PostMessage(CM_KEYS_NOT_FOUND, 0, 0);
    return false;
}

void CChildView::deleteAllKeys()
{
    m_engine.deleteAllKeys();
    m_bToDisplaySimilarClms = false;
    m_bXSimilarityComputed = false;
    m_vecSimilaritiesAcrossTables.clear();
    m_vecSimilaritiesAcrossTablesSorted.clear();
}

void CChildView::OnRButtonUp(UINT nFlags, CPoint point)
{
    if (M_CCell.x * M_CCell.y)
    {
        if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns)
        {
            if (m_engine.deleteKey(1, M_CCell.y) + m_engine.deleteKey(2, M_CCell.x) == 0)
            {
                m_engine.pushKey(M_CCell.y, M_CCell.x);
            }
            this->Invalidate();
        }
    }
    CWnd::OnRButtonUp(nFlags, point);
}

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
            m_engine.pushKey(k1, k2);
        }
    }
    return false;
}

int CChildView::getNumberOfPossibleKeys()
{
    return m_keyFinder.getNumberOfPossibleKeys();
}

void CChildView::OnUpdateKeySearchComplexity(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(true);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    if (!m_pCombo2)
    {
        m_pCombo2 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_KEY_SEARCH_COMPLEXITY));
        if (m_pCombo2)
        {
            m_pCombo2->SelectItem(1);
        }
    }
}

void CChildView::OnKeySearchComplexity()
{
    int complexity = 100000;
    if (m_pCombo2->GetCurSel() == 0)
        complexity = 10000;
    if (m_pCombo2->GetCurSel() == 1)
        complexity = 100000;
    if (m_pCombo2->GetCurSel() == 2)
        complexity = 1000000;
    m_keyFinder.setComplexity(complexity);
}

int CChildView::getNumberOfPossibleKeys(int table, int order, int item)
{
    return m_keyFinder.getNumberOfPossibleKeys(table, order, item);
}

void CChildView::findSims() // Fallback function: do not use in case there is sufficient RAM capacity
{
    COleVariant vData;
    CString cellContent; // Renamed from szdata
    long currentSimilarityScore; // Renamed from tmpSim
    int lastReportedProgress = 0, currentProgress = 0; // Renamed from prgHlpr0, prgHlpr
    
    // Reset the cross-table similarity containers
    m_vecSimilaritiesAcrossTables.clear();
    m_vecSimilaritiesAcrossTablesSorted.clear();
    
    // Initialize the vector with zeroed values up to the column count of Table 1
    SimilaritiesAcrossTables tempSimilarity;
    m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);
    for (int colIndex1 = 1; colIndex1 <= m_Table1.NumberOfColumns + 1; colIndex1++) // Renamed from tmp_i
    {
        tempSimilarity.similarityOrder = 0;
        tempSimilarity.similarity = 0;
        tempSimilarity.clm1 = colIndex1;
        tempSimilarity.clm2 = 0;
        m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);
    }
    
    // Compare each column in Table 1 against all columns in Table 2
    for (int colIndex1 = 1; colIndex1 <= m_Table1.NumberOfColumns; colIndex1++) // Renamed from c_i1
    {
        lastReportedProgress = 100 * colIndex1 / m_Table1.NumberOfColumns;
        if (lastReportedProgress > currentProgress)
        {
            currentProgress = lastReportedProgress;
            PostMessage(CM_UPDATE_KEYPROGRESS1, 0, currentProgress);
        }
        
        m_mapTmpMap1.clear();
        
        // Populate the frequency map for the current column in Table 1
        for (int row1 = m_Table1.FirstRowWithData; row1 <= m_Table1.NumberOfRows; row1++) // Renamed from r_i1
        {
            cellContent = m_excel1.getCellValue(colIndex1, row1);
            if ((cellContent != L"") && (m_mapTmpMap1.find(cellContent) == m_mapTmpMap1.end()))
            {
                m_mapTmpMap1[cellContent] = row1;
            }
        }
        
        // Check this frequency map against every column in Table 2
        for (int colIndex2 = 1; colIndex2 <= m_Table2.NumberOfColumns; colIndex2++) // Renamed from c_i2
        {
            currentSimilarityScore = 0;
            m_mapTmpMap2.clear();
            
            for (int row2 = m_Table2.FirstRowWithData; row2 <= m_Table2.NumberOfRows; row2++) // Renamed from r_i2
            {
                cellContent = m_excel2.getCellValue(colIndex2, row2);
                if ((cellContent != L"") && (m_mapTmpMap1.find(cellContent) != m_mapTmpMap1.end()))
                {
                    // If the content is found in both columns, count it as a similarity match
                    if (m_mapTmpMap2.find(cellContent) == m_mapTmpMap2.end())
                    {
                        m_mapTmpMap2[cellContent] = row2;
                        currentSimilarityScore++;
                    }
                }
            }
            
            // Record the highest similarity score for the current column in Table 1
            if (currentSimilarityScore > m_vecSimilaritiesAcrossTables[colIndex1].similarity)
            {
                m_vecSimilaritiesAcrossTables[colIndex1].similarity = currentSimilarityScore;
                m_vecSimilaritiesAcrossTables[colIndex1].clm1 = colIndex1;
                m_vecSimilaritiesAcrossTables[colIndex1].clm2 = colIndex2;
            }
        }
    }
    
    // Sort and rank the similarities to find the best overall column matches
    int similarityRank = 1; // Renamed from simOrder
    tempSimilarity.clm1 = 0;
    tempSimilarity.clm2 = 0;
    tempSimilarity.similarity = 0;
    tempSimilarity.similarityOrder = 0;
    m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
    
    for (int loopIndex = 1; loopIndex <= m_Table1.NumberOfColumns; loopIndex++) // Renamed from i0
    {
        tempSimilarity.clm1 = 0;
        tempSimilarity.clm2 = 0;
        tempSimilarity.similarity = 0;
        tempSimilarity.similarityOrder = 0;
        
        for (int colIndex1 = 1; colIndex1 <= m_Table1.NumberOfColumns; colIndex1++) // Renamed from i1
        {
            if (m_vecSimilaritiesAcrossTables[colIndex1].similarity > 0 &&
                m_vecSimilaritiesAcrossTables[colIndex1].similarity > tempSimilarity.similarity &&
                m_vecSimilaritiesAcrossTables[colIndex1].similarityOrder == 0) // Ensure we only rank unranked columns
            {
                tempSimilarity.similarityOrder = similarityRank;
                tempSimilarity.similarity = m_vecSimilaritiesAcrossTables[colIndex1].similarity;
                tempSimilarity.clm1 = m_vecSimilaritiesAcrossTables[colIndex1].clm1;
                tempSimilarity.clm2 = m_vecSimilaritiesAcrossTables[colIndex1].clm2;
            }
        }
        
        if (tempSimilarity.similarity > 0)
        {
            similarityRank++;
            m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
            m_vecSimilaritiesAcrossTables[tempSimilarity.clm1].similarityOrder = similarityRank;
        }
    }
    
    // Store the total number of matched columns at index 0
    m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder = similarityRank - 1; 
    
    PostMessage(CM_UPDATE_PROGRESS, 0, 0);
    this->Invalidate();
    
    // If we found at least one similarity, enable the similarity display functionality
    if (similarityRank > 1)
    {
        m_bToDisplaySimilarClms = true;
        m_bXSimilarityComputed = true;
    }
    
    m_bLockPrg1 = false;
    return;
}

void CChildView::findSimsRange(int startColIndex, int endColIndex, UINT progressMsg, UINT doneMsg, bool useTmp)
{
    CString cellContent; // Renamed from szdata
    long long currentSimilarityScore; // Renamed from tmpSim
    int lastReportedProgress = 0, currentProgress = 0; // Renamed from prgHlpr0, prgHlpr
    
    std::map<CString, long> table1FreqMap; // Renamed from thdSafe_tmpMap1
    std::map<CString, long> table2FreqMap; // Renamed from thdSafe_tmpMap2
    CString mapKey = L""; // Renamed from what
    long occurrenceCount1 = 0; // Renamed from occurence1
    long occurrenceCount2 = 0; // Renamed from occurence2
    long table1RowCount = 0; // Renamed from size1
    long table2RowCount = 0; // Renamed from size2
    long minRowCount = 0; // Renamed from minsize
    long maxRowCount = 0; // Renamed from maxsize
    double unitSimilarity = 0.f; // Renamed from tmpUnitSim
    long sizeRatio = 0; // Renamed from tmp_varRat
    long finalSimilarityScore = 0; // Renamed from sim
    long totalOccurrences1 = 0; // Renamed from sumOccurence1
    long totalOccurrences2 = 0; // Renamed from sumOccurence2
    long exactPureSimilarity; // Renamed from pureSim
    
    int rangeSize = endColIndex - startColIndex;
    rangeSize = rangeSize ? rangeSize : 1; // Prevent division by zero
    
    // Iterate over the specified range of columns in Table 1
    for (int colIndex1 = startColIndex; colIndex1 <= endColIndex; colIndex1++) // Renamed from c_i1
    {
        lastReportedProgress = 100 * (colIndex1 - startColIndex) / rangeSize;
        if (lastReportedProgress > currentProgress)
        {
            currentProgress = lastReportedProgress;
            PostMessage(progressMsg, 0, currentProgress);
        }
        
        table1FreqMap.clear();
        
        // Build the frequency map of strings for the current column in Table 1
        for (int row1 = m_Table1.FirstRowWithData; row1 <= m_Table1.NumberOfRows; row1++) // Renamed from r_i1
        {
            cellContent = useTmp ? m_excel1.getTmpCellValue(colIndex1, row1) : m_excel1.getCellValue(colIndex1, row1);
            if (cellContent != L"")
            {
                if (table1FreqMap.find(cellContent) == table1FreqMap.end())
                    table1FreqMap[cellContent] = 1;
                else
                    table1FreqMap[cellContent] = table1FreqMap[cellContent] + 1;
            }
        }
        
        // Compare the frequency map against all columns in Table 2
        for (int colIndex2 = 1; colIndex2 <= m_Table2.NumberOfColumns; colIndex2++) // Renamed from c_i2
        {
            table2FreqMap.clear();
            
            // Build the frequency map for the current column in Table 2, but only for strings that exist in the Table 1 map
            for (int row2 = m_Table2.FirstRowWithData; row2 <= m_Table2.NumberOfRows; row2++) // Renamed from r_i2
            {
                cellContent = useTmp ? m_excel2.getTmpCellValue(colIndex2, row2) : m_excel2.getCellValue(colIndex2, row2);
                if ((cellContent != L"") && (table1FreqMap.find(cellContent) != table1FreqMap.end()))
                {
                    if (table2FreqMap.find(cellContent) == table2FreqMap.end())
                        table2FreqMap[cellContent] = 1;
                    else
                        table2FreqMap[cellContent] = table2FreqMap[cellContent] + 1;
                }
            }
            
            totalOccurrences1 = totalOccurrences2 = 0;
            currentSimilarityScore = 0;
            
            // Calculate similarity penalty based on frequency mismatches between the two columns
            for (auto mapIterator : table1FreqMap) // Renamed from iterator
            {
                mapKey = mapIterator.first;
                occurrenceCount1 = mapIterator.second;
                totalOccurrences1 += occurrenceCount1;
                occurrenceCount2 = 0;
                
                if (table2FreqMap.find(mapKey) != table2FreqMap.end())
                {
                    occurrenceCount2 = table2FreqMap[mapKey];
                    totalOccurrences2 += occurrenceCount2;
                    // The difference in occurrence counts acts as a penalty
                    unitSimilarity = max(occurrenceCount1, occurrenceCount2) - min(occurrenceCount1, occurrenceCount2);
                    currentSimilarityScore += (unitSimilarity);
                }
            }
            
            finalSimilarityScore = currentSimilarityScore;
            table1RowCount = m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1;
            table2RowCount = m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1;
            minRowCount = min(table1RowCount, table2RowCount);
            minRowCount = minRowCount ? minRowCount : 1;
            maxRowCount = max(table1RowCount, table2RowCount);
            
            // Normalize the similarity score to account for differently sized tables
            {
                sizeRatio = min(table1FreqMap.size(), table2FreqMap.size());
                if (sizeRatio)
                {
                    sizeRatio = sizeRatio ? sizeRatio : 1;
                    sizeRatio = (minRowCount < sizeRatio ? 1 : minRowCount / sizeRatio);
                    finalSimilarityScore = (minRowCount - finalSimilarityScore) / sizeRatio + 1;
                }
                
                // If tables are identical in matched size and perfect match, force a score of 1
                if (finalSimilarityScore == 0 && table1FreqMap.size() == table2FreqMap.size() && sizeRatio > 0)
                {
                    finalSimilarityScore = 1;
                }
            }
            
            // Calculate the pure similarity score that takes into account the discrepancy in overall occurrences
            exactPureSimilarity = (maxRowCount - abs(totalOccurrences2 - totalOccurrences1)) - currentSimilarityScore;
            
            // Update the global similarity metrics if this column combination is better than previous ones
            if (exactPureSimilarity > m_vecSimilaritiesAcrossTables[colIndex1].pureSim && finalSimilarityScore > 0)
            {
                m_vecSimilaritiesAcrossTables[colIndex1].similarity = min(table1FreqMap.size(), table2FreqMap.size());
                m_vecSimilaritiesAcrossTables[colIndex1].clm1 = colIndex1;
                m_vecSimilaritiesAcrossTables[colIndex1].clm2 = colIndex2;
                m_vecSimilaritiesAcrossTables[colIndex1].pureSim = exactPureSimilarity;
            }
        }
    }
    
    // Notify thread that computation is complete
    PostMessage(doneMsg, 0, 0);
}

void CChildView::findSims1()
{
    int tmp_bnd_hlf = m_Table1.NumberOfColumns / 2;
    findSimsRange(1, tmp_bnd_hlf, CM_UPDATE_KEYPROGRESS1, CM_SIMS1_DONE, false);
}

void CChildView::findSims2()
{
    int tmp_bnd_hlf = m_Table1.NumberOfColumns / 2;
    findSimsRange(m_Table1.NumberOfColumns - tmp_bnd_hlf, m_Table1.NumberOfColumns, CM_UPDATE_KEYPROGRESS2,
                  CM_SIMS2_DONE, true);
}

void CChildView::OnShowSimilarColumns()
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

void CChildView::OnUpdateShowSimilarColumns(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(m_Table1.NumberOfColumns * m_Table2.NumberOfColumns && m_bXSimilarityComputed);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pShowSims = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_SHOW_SIMILAR_COLUMNS));
    pCmdUI->SetCheck(m_bToDisplaySimilarClms);
}

void CChildView::OnFindColumnRelations()
{
    if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns == 0)
    {
        return;
    }
    if (m_bLockPrg1 || m_bLockPrg2)
    {
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
        m_vecSimilaritiesAcrossTables.push_back(tempSimilarity);
        ;
    }
    // </Preparation for actual-relations check>
    m_bToDisplaySimilarClms = false;
    m_bXSimilarityComputed = false;
    AfxBeginThread(FindSimsThreadProc1, this);
    AfxBeginThread(FindSimsThreadProc2, this);
    m_bLockPrg1 = true;
    m_bLockPrg2 = true;
}

void CChildView::OnIdxcrtBtn() {}

afx_msg LRESULT CChildView::OnCmUpdateKeyProgress1(WPARAM wParam, LPARAM lParam)
{
    m_pKeyProgressBar1->SetPos((UINT)lParam);
    return 0;
}

afx_msg LRESULT CChildView::OnCmUpdateKeyProgress2(WPARAM wParam, LPARAM lParam)
{
    m_pKeyProgressBar2->SetPos((UINT)lParam);
    return 0;
}

afx_msg LRESULT CChildView::OnCmKeys1Done(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar1->SetPos(0);
    m_bLockPrg1 = false;
    this->Invalidate();
    if (m_bDoAutoMark)
    {
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
        resolveAutoMark();
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
    }
    if (m_bWaitingForKeys)
    {
        m_bKeys1done = true;
        if (m_bKeys2done)
        {
            m_bWaitingForKeys = false;
            m_bKeys1done = false;
            m_bKeys2done = false;
            AfxBeginThread(MyThreadProc, this);
            if (g_pMainFrame)
                g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
        }
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmKeys2Done(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar2->SetPos(0);
    m_bLockPrg2 = false;
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
            if (g_pMainFrame)
                g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
        }
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmGathering1Done(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar1->SetPos(0);
    m_bLockPrg1 = false;
    this->Invalidate();
    if (m_bDoAutoMark)
    {
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
        resolveAutoMark();
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
    }
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
            if (g_pMainFrame)
                g_pMainFrame->updateStatusBar(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING));
        }
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmGathering2Done(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar2->SetPos(0);
    m_bLockPrg2 = false;
    if (m_bWaitingForKeys)
    {
        m_bKeysGathering2done = true;
        if (m_bKeysGathering1done)
        {
            m_bWaitingForKeys = false;
            m_bKeysGathering1done = false;
            m_bKeysGathering2done = false;
            AfxBeginThread(MutualCheckThreadProc, this);
            if (g_pMainFrame)
                g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
        }
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmKeysFound(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar1->SetPos(0);
    m_bLockPrg1 = false;
    this->Invalidate();
    if (m_bDoAutoMark)
    {
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
        resolveAutoMark();
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
    }
    m_bWaitingForKeys = false;
    usePossibleKeys();
    CString tmpS;
    tmpS.Format(CMsg(IDS_KEY_COMB_FOUND), m_keyFinder.getBestKeyComb().cnt); // CMsg(IDS_KEY_COMB_FOUND)
    MessageBox(tmpS);
    EndWaitCursor();
    return 0;
}

afx_msg LRESULT CChildView::OnCmKeysNotFound(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar1->SetPos(0);
    m_bLockPrg1 = false;
    this->Invalidate();
    if (m_bDoAutoMark)
    {
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
        resolveAutoMark();
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
    }
    m_bWaitingForKeys = false;
    MessageBox(CMsg(IDS_INCOMPATBL_KEY_FOUND)); // CMsg(IDS_INCOMPATBL_KEY_FOUND)
    EndWaitCursor();
    return 0;
}

afx_msg LRESULT CChildView::OnCmSims1Done(WPARAM wParam, LPARAM lParam)
{
    m_pKeyProgressBar1->SetPos(0);
    m_bLockPrg1 = false;
    if (m_bLockPrg2 == false)
    {
        finishFindRelations();
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmSims2Done(WPARAM wParam, LPARAM lParam)
{
    m_pKeyProgressBar2->SetPos(0);
    m_bLockPrg2 = false;
    if (m_bLockPrg1 == false)
    {
        finishFindRelations();
    }
    return 0;
}

afx_msg LRESULT CChildView::OnCmFirstPassDone(WPARAM wParam, LPARAM lParam)
{
    m_pProgressBar1->SetPos(0);
    m_bLockPrg1 = false;
    this->Invalidate();
    if (m_bDoAutoMark)
    {
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
        resolveAutoMark();
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
    }
    if (m_nEffMax)
    {
        m_szRsltTxt.Format(CMsg(IDS_FOUND_KEYS_FROM_TOTAL), m_nEffMax,
                           (m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1),
                           (m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1)); // CMsg(IDS_FOUND_KEYS_FROM_TOTAL)
        if (g_pMainFrame)
            g_pMainFrame->updateStatusBar(m_szRsltTxt);
    }
    return 0;
}

void CChildView::OnUpdateKeyProgress1(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pKeyProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_KEY_PROGRESS1));
}

void CChildView::OnUpdateKeyProgress2(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pKeyProgressBar2 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_KEY_PROGRESS2));
}

void CChildView::finishFindRelations()
{
    if (m_bXSimilarityComputed)
    {
        return;
    }
    
    SimilaritiesAcrossTables tempSimilarity;
    int similarityRank = 1; // Renamed from simOrder
    tempSimilarity.clm1 = 0;
    tempSimilarity.clm2 = 0;
    tempSimilarity.similarity = 0;
    tempSimilarity.similarityOrder = 0;
    
    m_vecSimilaritiesAcrossTablesSorted.clear();
    m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity); // Index 0 will hold metadata
    
    // Sort and rank all columns by their similarity scores computed by background threads
    for (int loopIndex = 1; loopIndex <= m_Table1.NumberOfColumns; loopIndex++) // Renamed from i0
    {
        tempSimilarity.clm1 = 0;
        tempSimilarity.clm2 = 0;
        tempSimilarity.similarity = -1;
        tempSimilarity.similarityOrder = 0;
        
        for (int colIndex1 = 1; colIndex1 <= m_Table1.NumberOfColumns; colIndex1++) // Renamed from i1
        {
            if (m_vecSimilaritiesAcrossTables[colIndex1].similarity > tempSimilarity.similarity &&
                m_vecSimilaritiesAcrossTables[colIndex1].similarityOrder == 0) // Ensure column isn't already ranked
            {
                tempSimilarity.similarityOrder = similarityRank;
                tempSimilarity.similarity = m_vecSimilaritiesAcrossTables[colIndex1].similarity;
                tempSimilarity.clm1 = m_vecSimilaritiesAcrossTables[colIndex1].clm1;
                tempSimilarity.clm2 = m_vecSimilaritiesAcrossTables[colIndex1].clm2;
            }
        }
        
        // Push the highest available similarity into the sorted list
        {
            similarityRank++;
            m_vecSimilaritiesAcrossTablesSorted.push_back(tempSimilarity);
            m_vecSimilaritiesAcrossTables[tempSimilarity.clm1].similarityOrder = similarityRank;
        }
    }
    
    // At the zero position, store the total number of columns that have a "lookalike" in the second file
    m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder = similarityRank - 1;
    
    this->Invalidate();
    
    // If we successfully ranked columns, enable the similarity visualization feature
    if (similarityRank > 1)
    {
        m_bToDisplaySimilarClms = true;
        m_bXSimilarityComputed = true;
    }
}

void CChildView::OnUpdateIdxCheckbox(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_IDX_CHECKBOX));
    pCmdUI->SetCheck(m_bUseIndexes);
}

int CChildView::ReverseFind(LPCTSTR lpszData, LPCTSTR lpszSub, int startpos)
{
    int lenSub = lstrlen(lpszSub);
    int len = lstrlen(lpszData);
    if (0 < lenSub && 0 < len)
    {
        if (startpos == -1 || startpos >= len)
            startpos = len - 1;
        for (LPCTSTR lpszReverse = lpszData + startpos; lpszReverse != lpszData; --lpszReverse)
            if (_tcsncmp(lpszSub, lpszReverse, lenSub) == 0)
                return (lpszReverse - lpszData);
    }
    return -1;
}

void CChildView::OnCheckIdx()
{
    m_bUseIndexes = !m_bUseIndexes;
    m_engine.m_bUseIndexes = m_bUseIndexes;
}

void CChildView::OnUpdateCheckIdx(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_IDX_CHECKBOX));
    pCmdUI->SetCheck(m_bUseIndexes);
}

void CChildView::OnUseKeyIndexing()
{
    m_bUseIndexes = !m_bUseIndexes;
    m_engine.m_bUseIndexes = m_bUseIndexes;
    if (m_bUseIndexes)
        MessageBox(CMsg(IDS_IDXING_WARNING)); // CMsg(IDS_IDXING_WARNING)
}

void CChildView::OnUpdateUseKeyIndexing(CCmdUI* pCmdUI)
{
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_USE_KEY_INDEXING));
    pCmdUI->SetCheck(m_bUseIndexes);
}

void CChildView::OnUpdateRows1(CCmdUI* pCmdUI)
{
    if (!(m_szFilename1 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pRows1 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_ROWS1));
}

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

void CChildView::OnUpdateCols1(CCmdUI* pCmdUI)
{
    if (!(m_szFilename1 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pCols1 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_COLS1));
}

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

void CChildView::OnUpdateRows2(CCmdUI* pCmdUI)
{
    if (!(m_szFilename2 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pRows2 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_ROWS2));
}

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

void CChildView::OnUpdateCols2(CCmdUI* pCmdUI)
{
    if (!(m_szFilename2 == ""))
        pCmdUI->Enable(true);
    else
        pCmdUI->Enable(false);
    m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
    m_pCols2 = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_COLS2));
}

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
