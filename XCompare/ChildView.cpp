
#include "stdafx.h"
#include <cstring>
#include "XCompare.h"
#include "ChildView.h"
#include "MainFrm.h"
#include "Msg.h"

extern CMainFrame* g_pMainFrame; // pointer to FrameWindow

#define LINESIZE 8 // thickness of thick lines drawn in the visual representation of the result matrix

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define MAX_CFileDialog_FILE_COUNT 1 // Number of selectable files in file selectors
#define FILE_LIST_BUFFER_SIZE ((MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1) 

#define CM_UPDATE_PROGRESS WM_APP + 1
#define CM_UPDATE_PROGRESS2 WM_APP + 2
#define CM_UPDATE_PROGRESS3 WM_APP + 3

#define OFFSET_Y 100 // height of a column (within the visual representation of the result matrix) that contains the names of column taken from the second table
#define OFFSET_X 100 // width of a row ... first table
#define STEP_X 24 // width of a cell ...
#define STEP_Y 24 // height of a cell ...


BOOL uniqueKeys1; // to indicate whether values taken from key columns are identical - this is an inherent prerequisite for the comparison to be successful
BOOL uniqueKeys2; // the same as above - for the second file


bool waitingForKeys; // indicates status of waiting for maps of keys
bool keys1done; // indicates status of readiness of keys for the first table
bool keys2done; // indicates status of readiness of keys for the second table

CString rsltTxt; // human understandable text indicating some information related to the matrix - displayed in the status bar

struct Palette {
	int red;
	int green;
	int blue;
}; // structure type for color of cells

int threadCnt;

struct Table {
	int WorkSheetNumber;
	long NumberOfRows;
	int FirstRowWithData;
	int RowWithNames;
	int NumberOfColumns;
	CString Columns[255];
	int keys[3];
}; // structure type for description of tables

struct VisTopLeft {
	int top;
	int left;
}; // contains coords (in the units of matrix cells) of scrolled matrix


struct ChosenCell {
	int x;
	int y;
}; // contains coords (in the matrix cells) of the cell that is pointed by mouse

struct Clnt {
	int w;
	int h;
}; // Size of client area (in pixels)

struct NotUniqueKeys {
	long firstRow;
	long secondRow;
	CString keyString;
}; // if keys are found to not be unique, this structure contains the rows of the first found duplicate
NotUniqueKeys notUniqueKeys; 

//int threadCounter; // to be even more scallable in future
//int stepCounter;

Palette palette[20]; // user can choose one of the 20 colors that will be used for background of found difference 

bool prereq1valid, prereq2valid; // are prerequisities for execution of main process fulfilled

int oldx, oldy; // coords of last chosen cell

int chosenColor1, chosenColor2; // color will be used as a background in XLS file to mark difference

Clnt clnt; // client area
bool lockPrg1; // indicates status of computing in threads (inversely)
bool lockPrg2; // ... same as above

bool doAutoMark; // whether user selected the option for automatic marking in XLS files

int matrixDone; // result matrix is done and ready for follow-up analysis
int prereqDone; // prerequisities for main comparison process are fulfilled

bool markIdentCols;  // not used

bool sameNames; // user wants the cells that intersects columns with same names to be marked by thick border

int effMax; // counted up number of keys that were found in both files

char *mainArr1; // this 2D array contains first character of content of each cell taken from the first XLS file (first horizontally, then vertically)
char *mainArr2; // .... the second XLS file

bool *markIn1Arr; // this 2D array indicates whether a cell at its coordinates (see above) is to be marked in the first file
bool *markIn2Arr; // ... in the second file.

CString *keyArr11; // array of the strings found in the first key column in the first file
CString *keyArr12; // ... the second key ... the first file
CString *keyArr13; // ... the third key ... the first file
CString *keyArr21; // ... the first key ... the second file
CString *keyArr22; // ... the second key ... the second file
CString *keyArr23; // ... the third key ... the second file

bool *keyMissing1;
bool *keyMissing2;

int *mainMatrix; // 2D array representing the result matrix
bool *markedMatrix; // 2D array indicating marked cells in the result matrix
bool *emptyClms1; // 1D array indicating empty columns in the first file
bool *emptyClms2; // .... the second file
bool *greenClms1; // 1D array indicating whether a column has its "lookalike" in the second file
bool *greenClms2; // 1D array indicating whether a column has its "lookalike" in the first file

long *foundDifferences; // number of differences found between intersected columns (for the doubleclicked cell)
long selectedDifference; // difference picked by user in the drop down box in the "analysis" tab

CMFCRibbonBar* pRibbon; // pointer to ribbon object
//CMFCRibbonStatusBarPane *statusBarPane;

Table table1; 
Table table2;

COleSafeArray saRet1; // OLE object for connection to first Excel file
COleSafeArray saRet2; // ... second Excel file

CString filename1; // name of the first file that is to be compared
CString filename2; // .... second ....

CWorkbooks books1; // TypeLib objects
CWorkbook book1; // ...
CWorksheets sheets1; // ...
CWorksheet sheet1; // ...
CRange oRange1; // ...


CWorkbooks books2; // TypeLib objects
CWorkbook book2; // ...
CWorksheets sheets2; // ...
CWorksheet sheet2; // ...
CRange oRange2; // ...

CCellFormat cellFormat; // TypeLib object
Cnterior interior; // TypeLib object

COleVariant
covTrue((short)TRUE),
covFalse((short)FALSE),
covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR); // OLE constants

CApplication app; // application object

CMap <CString, LPCTSTR, long, long> map1; // map for keys in the first file
CMap <CString, LPCTSTR, long, long> map2; // ... second file


int nUiToBeRefreshed; // how many times the UI is to be refreshed (just a workaround)
float fZoom; // not used at the moment
int prgval1; // not used 
CMFCRibbonProgressBar* pProgressBar1; // CMFCRibbon UI objects
CMFCRibbonProgressBar* pProgressBar2;

CMFCRibbonComboBox* pSheetCombo1;
CMFCRibbonComboBox* pSheetCombo2;
CMFCRibbonEdit* pSpinner1_Fdata;
CMFCRibbonEdit* pSpinner1_Names;
CMFCRibbonEdit* pSpinner2_Fdata;
CMFCRibbonEdit* pSpinner2_Names;
CMFCRibbonComboBox* pKeyCombo11;
CMFCRibbonComboBox* pKeyCombo12;
CMFCRibbonComboBox* pKeyCombo13;
CMFCRibbonComboBox* pKeyCombo21;
CMFCRibbonComboBox* pKeyCombo22;
CMFCRibbonComboBox* pKeyCombo23;
CMFCRibbonCheckBox* pMarkIn1;
CMFCRibbonCheckBox* pMarkIn2;
CMFCRibbonSlider* pSlider;
CMFCRibbonButton* pUnhideExcel;
CMFCRibbonCheckBox* pVerifyKeys;
CMFCRibbonCheckBox* pSameNames;
CMFCRibbonColorButton* pColorPicker1;
CMFCRibbonColorButton* pColorPicker2;
CMFCRibbonCheckBox* pAuto;
CMFCRibbonComboBox* pFoundDifferences;
CMFCRibbonLabel* pLabel0;
CMFCRibbonLabel* pLabel1;
CMFCRibbonLabel* pLabel2;
CMFCRibbonCheckBox* pToFront;


bool toFront; // should Excel be moved to front when the difference is requested to be shown?

int scrolled_X; // how many cells did we scroll horizontally?
int scrolled_Y; // how many cells did we scroll vertically?
ChosenCell cCell; // this structure contains coordinates of the cell the mouse pointer is hovering above.
VisTopLeft visTopLeft; // the coordinates of the topmost and leftmost visible cell

bool in1file; // whether are differences to be marked in the first file
bool in2file; // ... second file

bool autoMark; // do we request automatic marking of differences?

bool verifyKeys; // not used anymore - the check of the uniqueness of keys is mandatory and as such it is accomplished automatically

bool toInitSB; // not used anymore

int m_nCellWidth;   // Cell width in pixels
int m_nCellHeight;  // Cell height in pixels
int m_nRibbonWidth; // Ribbon width in pixels
int m_nViewWidth;   // Workspace width in pixels
int m_nViewHeight;  // Workspace height in pixels
int m_nHScrollPos;  // Horizontal scroll position
int m_nVScrollPos;  // Vertical scroll position
int m_nHPageSize;   // Horizontal page size
int m_nVPageSize;   // Vertical page size

bool onlyPcnt; // whether we want to see detail of a hovered cell
bool forceNotOnlyPcnt; // inverse of the above (just a helper)

int sldr; // value set on the slider in the "analysis" tab

// CChildView

// <Declaration of threads>

UINT MyThreadProc(LPVOID pParam);
UINT MyThreadProc2(LPVOID pParam);
// CreateKeys1ThreadProc(LPVOID pParam);
UINT MyThreadProc3(LPVOID pParam);
UINT CreateKeys1ThreadProc(LPVOID pParam);
UINT CreateKeys2ThreadProc(LPVOID pParam);
UINT makePrereq1ThreadProc(LPVOID pParam);
UINT makePrereq2ThreadProc(LPVOID pParam);

CString mszRsrcs;

// </Declaration of threads>

CChildView::CChildView()
{
	threadCnt = 1;
	mszRsrcs = L"";
	toFront = false;
	selectedDifference = 0;
	forceNotOnlyPcnt = true;
	prereq1valid = false;
	prereq2valid = false;
	chosenColor1 = 13;
	chosenColor2 = 13;
	palette[0] = { 0,   0,   0 };
	palette[1] = { 128,   0,   0 };
	palette[2] = { 0,   128,   0 };
	palette[3] = { 128, 128,   0 };
	palette[4] = { 0,   0,   128 };
	palette[5] = { 128,   0, 128 };
	palette[6] = { 0,   128, 128 };
	palette[7] = { 192, 192, 192 };
	palette[8] = { 192, 220, 192 };
	palette[9] = { 166, 202, 240 };
	palette[10] = { 255, 251, 240 };
	palette[11] = { 160, 160, 164 };
	palette[12] = { 128, 128, 128 };
	palette[13] = { 255,   0,   0 };
	palette[14] = { 0,   255,   0 };
	palette[15] = { 255, 255,   0 };
	palette[16] = { 0,   0,   255 };
	palette[17] = { 255,   0, 255 };
	palette[18] = {   0, 255, 255 };
	palette[19] = { 255, 255, 255 };

	autoMark = false;
	doAutoMark = false;
	rsltTxt = "";
	uniqueKeys1 = false;
	uniqueKeys2 = false;
	lockPrg1 = false;
	lockPrg2 = false;
	verifyKeys = false;
	filename1 = "";
	filename2 = "";
	nUiToBeRefreshed = 3;
	fZoom = 100;
	prgval1 = 100; // just for test
	//CMFCRibbonBar* pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();


	mainArr1 = new char[1]; // this 2D array contains first character of content of each cell taken from the first XLS file (first horizontally, then vertically)
	mainArr2 = new char[1]; // .... the second XLS file

	markIn1Arr = new bool[1]; // this 2D array indicates whether a cell at its coordinates (see above) is to be marked in the first file
	markIn2Arr = new bool[1]; // ... in the second file.

	keyArr11 = new CString[1]; // array of the strings found in the first key column in the first file
	keyArr12 = new CString[1]; // ... the second key ... the first file
	keyArr13 = new CString[1]; // ... the third key ... the first file
	keyArr21 = new CString[1]; // ... the first key ... the second file
	keyArr22 = new CString[1]; // ... the second key ... the second file
	keyArr23 = new CString[1]; // ... the third key ... the second file

	keyMissing1 = new bool[1];
	keyMissing2 = new bool[1];

	mainMatrix = new int[1]; // 2D array representing the result matrix
	markedMatrix = new bool[1]; // 2D array indicating marked cells in the result matrix
	emptyClms1 = new bool[1]; // 1D array indicating empty columns in the first file
	emptyClms2 = new bool[1]; // .... the second file
	greenClms1 = new bool[1]; // 1D array indicating whether a column has its "lookalike" in the second file
	greenClms2 = new bool[1]; // 1D array indicating whether a column has its "lookalike" in the first file

	foundDifferences = new long[1]; // number of differences found between intersected columns (for the doubleclicked cell) 



	in1file = false;
	in2file = false;

	sameNames = false;

	cCell.x = 0;
	cCell.y = 0;

	onlyPcnt = false;


	toInitSB = true;

	visTopLeft.left = 0;
	visTopLeft.top = 0;


	sldr = 90;

	effMax = 0;

}

CChildView::~CChildView()
{
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
	ON_UPDATE_COMMAND_UI(ID_KEY1_1, &CChildView::OnUpdateKey11)
	ON_UPDATE_COMMAND_UI(ID_KEY1_2, &CChildView::OnUpdateKey12)
	ON_UPDATE_COMMAND_UI(ID_KEY1_3, &CChildView::OnUpdateKey13)
	ON_COMMAND(ID_KEY1_1, &CChildView::OnKey11)
	ON_COMMAND(ID_KEY1_2, &CChildView::OnKey12)
	ON_COMMAND(ID_KEY1_3, &CChildView::OnKey13)
	ON_COMMAND(ID_PICK_SECOND_SHEET, &CChildView::OnPickSecondSheet)
	ON_UPDATE_COMMAND_UI(ID_SPIN2_FDATA, &CChildView::OnUpdateSpin2Fdata)
	ON_COMMAND(ID_SPIN2_FDATA, &CChildView::OnSpin2Fdata)
	ON_UPDATE_COMMAND_UI(ID_SPIN2_NAMES, &CChildView::OnUpdateSpin2Names)
	ON_COMMAND(ID_SPIN2_NAMES, &CChildView::OnSpin2Names)
	ON_COMMAND(ID_KEY2_1, &CChildView::OnKey21)
	ON_UPDATE_COMMAND_UI(ID_KEY2_1, &CChildView::OnUpdateKey21)
	ON_UPDATE_COMMAND_UI(ID_KEY2_2, &CChildView::OnUpdateKey22)
	ON_UPDATE_COMMAND_UI(ID_KEY2_3, &CChildView::OnUpdateKey23)
	ON_COMMAND(ID_KEY2_2, &CChildView::OnKey22)
	ON_COMMAND(ID_KEY2_3, &CChildView::OnKey23)
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
END_MESSAGE_MAP()



// CChildView message handlers

BOOL CChildView::PreCreateWindow(CREATESTRUCT& cs) 
{
	if (!CWnd::PreCreateWindow(cs))
		return FALSE;

	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;
	cs.style |= WS_VSCROLL | WS_HSCROLL;
	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS, 
		::LoadCursor(NULL, IDC_ARROW), reinterpret_cast<HBRUSH>(COLOR_WINDOW+1), NULL);

	pRibbon = NULL;

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
	clnt.w = (rect.Width());
	clnt.h = (rect.Height());


	CPen pen1, pen2, pen3, pen4, pen5, pen6, pen7, pen8, pen9;
	CBrush brush0, brush1, brush2, brush3, brush4;

	pen1.CreatePen(PS_SOLID, 1, RGB(220, 220, 220));
	pen2.CreatePen(PS_SOLID, 1, RGB(200, 200, 200));
	pen3.CreatePen(PS_SOLID, 3, RGB(0, 0, 0));
	pen4.CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
	pen5.CreatePen(PS_SOLID, 1, RGB(100, 255, 100));
	pen6.CreatePen(PS_SOLID, 1, RGB(255, 255, 0));
	pen7.CreatePen(PS_SOLID, 2, RGB(0, 255, 0));
	pen8.CreatePen(PS_SOLID, 3, RGB(255, 0, 0));
	pen9.CreatePen(PS_SOLID, 2, RGB(155, 155, 255));

	brush0.CreateSolidBrush(RGB(255, 255, 255));
	brush1.CreateSolidBrush(RGB(100, 255, 100));
	brush2.CreateSolidBrush(RGB(255, 255, 0));
	brush3.CreateSolidBrush(RGB(255, 127, 50));
	brush4.CreateSolidBrush(RGB(255, 80, 80));
		

	dc.SelectObject(&pen2);
	dc.SelectObject(&brush1);


	CFont font1, font2, font3, font4;

	font1.CreateFontW(16, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font2.CreateFontW(16, 0, 900, 900, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font3.CreateFontW(12, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
	font4.CreateFontW(30, 0, 0, 0, 400, FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");



	int bnd_X_min = 1;
	int bnd_X_max = table2.NumberOfColumns;
	int bnd_Y_min = 1;
	int bnd_Y_max = table1.NumberOfColumns; // ?????

	//visTopLeft.left = 3;



	//dc.TextOutW(100, 100, L"+ěščřžýá"); // test of the ability to make an output in czech lang.
	dc.SelectObject(&brush0);

	int mx_x_adj, mx_y_adj; // for access to data

	dc.SelectObject(&pen2);
	dc.SelectObject(&brush1);

	for (int i_0 = OFFSET_X + STEP_X; i_0 < clnt.w; i_0 += STEP_X)
	{
		dc.MoveTo(i_0, 0);
		dc.LineTo(i_0, clnt.h);

	}
	for (int i_0 = OFFSET_Y + STEP_Y; i_0 < clnt.h; i_0 += STEP_X)
	{
		dc.MoveTo(0, i_0);
		dc.LineTo(clnt.w, i_0);
	}

	if (matrixDone && !onlyPcnt)
	{
		for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max; mx_y++)
		{
			if (((mx_y - visTopLeft.top) > 0))
			{
				if (greenClms1[mx_y])
				{
					dc.SelectObject(&pen5);
					dc.SelectObject(&brush1);
					dc.Ellipse(OFFSET_X, OFFSET_X + (mx_y - visTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1, OFFSET_Y + STEP_Y + (mx_y - visTopLeft.top) * STEP_Y);

				}
				if (emptyClms1[mx_y])
				{
					dc.SelectObject(&pen6);
					dc.SelectObject(&brush2);
					dc.Ellipse(OFFSET_X, OFFSET_X + (mx_y - visTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1, OFFSET_Y + STEP_Y + (mx_y - visTopLeft.top) * STEP_Y);
				}
			}

		}
		for (int mx_x = bnd_X_min; mx_x <= bnd_X_max; mx_x++)
		{
			if (((mx_x - visTopLeft.left) > 0))
			{
				if (greenClms2[mx_x])
				{
					dc.SelectObject(&pen5);
					dc.SelectObject(&brush1);
					dc.Ellipse(OFFSET_X + (mx_x - visTopLeft.left) * STEP_X, OFFSET_Y, OFFSET_X + STEP_X + (mx_x - visTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
				}
				if (emptyClms2[mx_x])
				{
					dc.SelectObject(&pen6);
					dc.SelectObject(&brush2);
					dc.Ellipse(OFFSET_X + (mx_x - visTopLeft.left) * STEP_X, OFFSET_Y, OFFSET_X + STEP_X + (mx_x - visTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
				}
			}
		}
	}


	dc.SelectObject(&pen2);
	dc.SelectObject(&brush1);



	dc.SelectObject(&font1);
	for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max; mx_y++)
	{
		mx_y_adj = mx_y + visTopLeft.top;
		dc.SetBkMode(OPAQUE);

		dc.Rectangle(0, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X, 1 + OFFSET_Y + mx_y * STEP_Y);
		dc.SetBkMode(TRANSPARENT);
		dc.TextOutW(2, OFFSET_Y + 5 + mx_y * STEP_Y, table1.Columns[mx_y_adj]);
	}
	dc.SelectObject(&font2);
	for (int mx_x = bnd_X_min; mx_x <= bnd_X_max; mx_x++)
	{
		mx_x_adj = mx_x + visTopLeft.left;
		dc.SetBkMode(OPAQUE);


		dc.Rectangle(OFFSET_X + mx_x * STEP_X, 0, 1 + OFFSET_X + mx_x * STEP_X, 1 + OFFSET_Y + STEP_Y );
		dc.SetBkMode(TRANSPARENT);
		dc.TextOutW(OFFSET_X + 5 + mx_x * STEP_X, -2 + OFFSET_Y + STEP_Y, table2.Columns[mx_x_adj]);
	}


	if (matrixDone && !onlyPcnt)
	{
		dc.SetBkMode(OPAQUE);
		dc.SelectObject(&font3);
		int valSimil;
		CString strSimil;
		if (effMax)
		{
			for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max - visTopLeft.top; mx_y++)
			{
				for (int mx_x = bnd_X_min; mx_x <= bnd_X_max - visTopLeft.left; mx_x++)
				{

					mx_y_adj = mx_y + visTopLeft.top;
					mx_x_adj = mx_x + visTopLeft.left;
					valSimil = mxGet(mx_x_adj, mx_y_adj) * 100 / effMax;
					strSimil.Format(L"%i", valSimil);
					strSimil += L"%";

					dc.SetBkMode(OPAQUE);



					if (!sameNames || (table1.Columns[mx_y_adj] == table2.Columns[mx_x_adj]))
					{
						if (valSimil == 100)
						{
							if (emptyClms1[mx_y_adj] || emptyClms2[mx_x_adj])
							{
								dc.SelectObject(&brush2);
							}
							else
							{
								dc.SelectObject(&brush1);
							}
						}
 
						if (valSimil < 100)
						{
							if (valSimil > sldr)
							{
								dc.SelectObject(&brush4);
							}
							else
							{
								dc.SelectObject(&brush0);
							}

						}
					}
					else
					{
						dc.SelectObject(&brush0);
					}


					dc.Rectangle(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X + mx_x * STEP_X, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y);
					dc.SetBkMode(TRANSPARENT);
					if (mxMarkedGet(mx_x_adj, mx_y_adj))
					{
						dc.SelectObject(&pen4);
						dc.MoveTo(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y);
						dc.LineTo(OFFSET_X + STEP_X + mx_x * STEP_X, OFFSET_Y + STEP_Y + mx_y * STEP_Y);
						dc.MoveTo(OFFSET_X + STEP_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y);
						dc.LineTo(OFFSET_X + mx_x * STEP_X, OFFSET_Y + STEP_Y + mx_y * STEP_Y);
						dc.SelectObject(&pen2);
					}
					dc.TextOutW(OFFSET_X + mx_x * STEP_X + 1, OFFSET_Y + mx_y * STEP_Y + 7, strSimil);




				}
			}


				dc.SetBkMode(TRANSPARENT);
				dc.SelectObject(GetStockObject(NULL_BRUSH));
				dc.SelectObject(&pen3);
				for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max - visTopLeft.top; mx_y++)
				{
					for (int mx_x = bnd_X_min; mx_x <= bnd_X_max - visTopLeft.left; mx_x++)
					{

						mx_y_adj = mx_y + visTopLeft.top;
						mx_x_adj = mx_x + visTopLeft.left;

						if (table1.Columns[mx_y_adj] == table2.Columns[mx_x_adj])
						{
							
							dc.Rectangle(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X + mx_x * STEP_X, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y);

						}

						if (mx_y_adj == oldy && mx_x_adj == oldx)
						{
							dc.SelectObject(&pen9);
							dc.Rectangle(OFFSET_X + mx_x * STEP_X + 3, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y - 4, 1 + OFFSET_X + STEP_X + mx_x * STEP_X - 2, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y - 2);
							dc.SelectObject(&pen3);
						}

					}

				}
				dc.SelectObject(&pen2);

		}
		onlyPcnt = false;

	}

	dc.SelectObject(&pen4);
	dc.MoveTo(0, OFFSET_Y + STEP_Y);
	dc.LineTo(clnt.w, OFFSET_Y+STEP_Y);
	dc.MoveTo(OFFSET_X+STEP_X, 0);
	dc.LineTo(OFFSET_X+STEP_X, clnt.h);

	dc.SelectObject(&font4);
	CString prcnt;
	prcnt = L"";
	if (cCell.x * cCell.y && matrixDone)
	{
		dc.SetBkMode(TRANSPARENT);
		if (cCell.x <= bnd_X_max && cCell.y <= bnd_Y_max)
		{
			long sameness = mxGet(cCell.x, cCell.y);
			//dc.SelectObject(&pen8);
			if (table1.Columns[cCell.y] == table2.Columns[cCell.x] && sameness < effMax)
				dc.SetTextColor(RGB(255, 0, 0));
			else
				dc.SetTextColor(RGB(0, 0, 0));

			prcnt.Format(L"Δ:%i", effMax - sameness);
			dc.TextOutW(5, 20, prcnt);
			dc.SetTextColor(RGB(0, 255, 0));

			prcnt.Format(L"=:%i", sameness);
			dc.TextOutW(5, 50, prcnt);
			if (emptyClms1[cCell.y] && emptyClms2[cCell.x])
			{
				dc.SetTextColor(RGB(0, 0, 0));
				dc.TextOutW(5, 80, CMsg(IDS_EMPTY)); // CMsg(IDS_EMPTY)
			}
		}

	}
	onlyPcnt = false;

	
}



void CChildView::OnPickFirstFile()
{

	g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_TILL_IN_EXCEL)); // CMsg(IDS_WAIT_TILL_IN_EXCEL)


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


		if (book1)
		{
			try {
				book1.Close(covOptional, covOptional, covOptional);
			}
			catch (COleException* e)
			{

			}
			
		}



	if (!(CString(fileName) == L""))
	{
		for (int i = 0; i < 255; i++)
		{
			table1.Columns[i] = "";
		}
		if (sheets1 = GetWorksheets1(fileName))
		{

			pSheetCombo1->RemoveAllItems();

			for (int i = 1; i <= sheets1.get_Count(); i++)
			{
				if (CWorksheet tempSheet = sheets1.get_Item(COleVariant((short)i)))
				{
					pSheetCombo1->AddItem(tempSheet.get_Name());

				}
				else
				{
					break;
				}
			}
		}
	}

	filename1 = fileName;
	nUiToBeRefreshed = 3;
	if (matrixDone > 0)
	{
		mxClear(table2.NumberOfColumns + 1, table1.NumberOfColumns + 1);
		matrixDone = 0;
	}
	g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED)); // CMsg(IDS_FILE_SUCCESFULLY_LOADED)
	this->Invalidate();
}


void CChildView::OnPickSecondFile()
{

	g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_TILL_IN_EXCEL)); // // CMsg(IDS_WAIT_TILL_IN_EXCEL)


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

	if (book2)
	{
		try {
			book2.Close(covOptional, covOptional, covOptional);
		}
		catch (COleException* e)
		{

		}

	}

	if (!(CString(fileName) == L""))
	{
		for (int i = 0; i < 255; i++)
		{
			table2.Columns[i] = "";
		}
		if (sheets2 = GetWorksheets2(fileName))
		{

			pSheetCombo2->RemoveAllItems();

			for (int i = 1; i <= sheets2.get_Count(); i++)
			{
				if (CWorksheet tempSheet = sheets2.get_Item(COleVariant((short)i)))
				{
					pSheetCombo2->AddItem(tempSheet.get_Name());
				}
				else
				{
					break;
				}
			}
		}
	}

	filename2 = fileName;
	nUiToBeRefreshed = 3;
	if (matrixDone > 0)
	{
		mxClear(table2.NumberOfColumns + 1, table1.NumberOfColumns + 1);
		matrixDone = 0;
	}
	g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED)); // CMsg(IDS_FILE_SUCCESFULLY_LOADED)
	this->Invalidate();
	
}


void CChildView::OnCreateMatrix()
{

	if (lockPrg1 || lockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}

	matrixDone = 0;
	prereqDone = 0;
	if ((table1.keys[0] + table1.keys[1] + table1.keys[2]) * (table2.keys[0] + table2.keys[1] + table2.keys[2]) == 0)
	{
		MessageBox(CMsg(IDS_ATLEAST_ONE_KEY)); // CMsg(IDS_ATLEAST_ONE_KEY)
		return;
	}

	g_pMainFrame->updateStatusBar(CMsg(IDS_COMPARISON_IN_PROGRESS)); // CMsg(IDS_COMPARISON_IN_PROGRESS)

	waitingForKeys = true;
	keys1done = false;
	keys2done = false;


	HWND hWnd0 = this->GetSafeHwnd();
	threadCnt++; AfxBeginThread(CreateKeys1ThreadProc, hWnd0);
	threadCnt++; AfxBeginThread(CreateKeys2ThreadProc, hWnd0);

	this->Invalidate();
	app.put_Visible(true);
	app.put_UserControl(TRUE);
	
}


void CChildView::OnUpdatePickFirstSheet(CCmdUI *pCmdUI)
{
	
		if (!(filename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false); 
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSheetCombo1 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_PICK_FIRST_SHEET));
}


void CChildView::OnUpdateCreateMatrix(CCmdUI *pCmdUI)
{

	if (nUiToBeRefreshed)
	{
		pCmdUI->Enable(true);
		if (!(filename1 == ""))
		{
			CMFCRibbonBar* pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
			pRibbon->ForceRecalcLayout();
			this->GetTopLevelFrame()->Invalidate();
		}

		if (!(filename2 == ""))
		{
			CMFCRibbonBar* pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
			pRibbon->ForceRecalcLayout();
			this->GetTopLevelFrame()->Invalidate();
		}
		if (nUiToBeRefreshed > 0 ) nUiToBeRefreshed  -= 1;
	}
}


void CChildView::OnUpdateFilename1(CCmdUI *pCmdUI)
{
	
	if (nUiToBeRefreshed)
	{

		if (!(filename1 == ""))
		{
			CString s = filename1;
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
		if (nUiToBeRefreshed > 0) nUiToBeRefreshed -= 1;
	}
}


void CChildView::OnUpdateFilename2(CCmdUI *pCmdUI)
{
	
	if (nUiToBeRefreshed)
	{
		if (!(filename2 == ""))
		{
			CString s = filename2;
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
		if (nUiToBeRefreshed > 0) nUiToBeRefreshed -= 1;
	}
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

	if (nDelta != 0) {
		m_nVScrollPos += nDelta;
		visTopLeft.top = m_nVScrollPos / STEP_Y;
		SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);
		//rect.left = OFFSET_X + STEP_X;
		rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(0, -nDelta, &rect);
		this->Invalidate();
	}

	onlyPcnt = false;
	forceNotOnlyPcnt = true;

	return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}


void CChildView::OnUpdatePickSecondSheet(CCmdUI *pCmdUI)
{
	
	if (!(filename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSheetCombo2 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_PICK_SECOND_SHEET));

}



void CChildView::OnUpdateProgress1(CCmdUI *pCmdUI)
{
	
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, pRibbon->FindByID(ID_PROGRESS2));


}


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
		visTopLeft.top = m_nVScrollPos / STEP_Y;
		SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);
		//rect.left = OFFSET_X + STEP_X;
		rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(0, -nDelta, &rect);
		onlyPcnt = false;
		forceNotOnlyPcnt = true;
		this->Invalidate();
	}

}


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
		visTopLeft.left = m_nHScrollPos / STEP_X;
		SetScrollPos(SB_HORZ, m_nHScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);
		rect.left = OFFSET_X + STEP_X;
		//rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(-nDelta, 0, &rect);
		this->Invalidate();
	}
}


CWorksheets CChildView::GetWorksheets1(CString TempBookName)
{
	if (!app)
	{
		if (!app.CreateDispatch(TEXT("Excel.Application")))
		{
			AfxMessageBox(CMsg(IDS_EXCEL_CANNOT_RUN)); // CMsg(IDS_EXCEL_CANNOT_RUN)
			return NULL;
		}
	}

	books1 = app.get_Workbooks();
	book1 = books1.Open(TempBookName, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);

	app.put_Visible(TRUE);
	app.put_UserControl(TRUE);

	return book1.get_Worksheets();
}

CWorksheets CChildView::GetWorksheets2(CString TempBookName)
{
	if (!app)
	{
		if (!app.CreateDispatch(TEXT("Excel.Application")))
		{
			AfxMessageBox(CMsg(IDS_EXCEL_CANNOT_RUN)); // CMsg(IDS_EXCEL_CANNOT_RUN)
			return NULL;
		}
	}

	books2 = app.get_Workbooks();
	book2 = books2.Open(TempBookName, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);

	app.put_Visible(TRUE);
	app.put_UserControl(TRUE);

	return book2.get_Worksheets();
}


void CChildView::OnPickFirstSheet()
{
	int tmpWSN = pSheetCombo1->GetCurSel() + 1;
	CString tmpWSS = pSheetCombo1->GetEditText();
	g_pMainFrame->updateStatusBar(L"Počkejte dokud nebude dokončena předběžná kontrola dat"); // CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)
	if (tmpWSN > 0)
	{
		saRet1.Destroy();
		table1.WorkSheetNumber = tmpWSN;
		sheet1 = sheets1.get_Item(COleVariant(tmpWSS));

		oRange1 = sheet1.get_UsedRange();

		saRet1 = oRange1.get_Value(covOptional);
		long iRows;
		long iCols;
		saRet1.GetUBound(1, &iRows);
		saRet1.GetUBound(2, &iCols);
		table1.NumberOfColumns = iCols;
		table1.NumberOfRows = iRows;
		table1.RowWithNames = 1;

		CString tmps;
		tmps.Format(_T("%d"), 1);
		pSpinner1_Names->SetEditText(tmps);
		table1.RowWithNames = 1;

		tmps.Format(_T("%d"), 2);
		pSpinner1_Fdata->SetEditText(tmps);
		table1.FirstRowWithData = 2;

		updateCombos1();

		m_nCellWidth = STEP_X;
		m_nCellHeight = STEP_Y;
		m_nRibbonWidth = 0;
		m_nViewWidth = STEP_X + OFFSET_X + ((table2.NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
		m_nViewHeight = STEP_Y + OFFSET_Y + m_nCellHeight * (table1.NumberOfColumns + 1);

		SCROLLINFO si;
		si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
		si.nMin = 0;
		si.nMax = m_nViewHeight - 1;
		si.nPos = m_nVScrollPos;
		si.nPage = m_nVPageSize;

		SetScrollInfo(SB_VERT, &si, TRUE);


		this->Invalidate();
		matrixDone = false;

		table1.keys[0] = 0;
		table1.keys[1] = 0;
		table1.keys[2] = 0;

		if (matrixDone > 0)
		{
			mxClear(table2.NumberOfColumns + 1, table1.NumberOfColumns + 1);
			matrixDone = 0;
		}

		HWND hWnd0 = this->GetSafeHwnd();
		g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED)); // CMsg(IDS_DATA_VERIFIED)
		threadCnt++; AfxBeginThread(makePrereq1ThreadProc, hWnd0);
	}
}


void CChildView::OnSpin1Names()
{
	CString tmps = pSpinner1_Names->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 1) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	pSpinner1_Names->SetEditText(tmps);

	table1.RowWithNames = tmpi;

	updateCombos1();
	this->Invalidate();

	
}


void CChildView::OnUpdateSpin1Names(CCmdUI *pCmdUI)
{
	
	if (!(filename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSpinner1_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, pRibbon->FindByID(ID_SPIN1_NAMES));
}


void CChildView::OnUpdateSpin1Fdata(CCmdUI *pCmdUI)
{
	
	if (!(filename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSpinner1_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, pRibbon->FindByID(ID_SPIN1_FDATA));
}


void CChildView::OnSpin1Fdata()
{
	CString tmps = pSpinner1_Fdata->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 2) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	pSpinner1_Fdata->SetEditText(tmps);
	table1.FirstRowWithData = tmpi;
	prereq1valid = false;
}


void CChildView::OnUpdateKey11(CCmdUI *pCmdUI)
{
	
	if (!(filename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pKeyCombo11 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_KEY1_1));
}


void CChildView::OnUpdateKey12(CCmdUI *pCmdUI)
{
	
	if (!(filename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pKeyCombo12 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_KEY1_2));
}


void CChildView::OnUpdateKey13(CCmdUI *pCmdUI)
{
	
	if (!(filename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pKeyCombo13 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_KEY1_3));
}


void CChildView::updateCombos1()
{
	pKeyCombo11->RemoveAllItems();
	pKeyCombo12->RemoveAllItems();
	pKeyCombo13->RemoveAllItems();
	pKeyCombo11->SetEditText(_T(""));
	pKeyCombo12->SetEditText(_T(""));
	pKeyCombo13->SetEditText(_T(""));
	long index[2];
	CString szdata;
	COleVariant vData;
	for (int i = 1; i <= table1.NumberOfColumns; i++)
	{
		// Loop through the data and report the contents.
		index[0] = table1.RowWithNames;
		index[1] = i;
		try {
			saRet1.GetElement(index, vData); vData = (CString)vData;
		}
		catch (COleException* e)
		{
			vData = L"";
		}
		szdata = vData;
		if (szdata == "") szdata = CMsg(IDS_NO_NAME); // CMsg(IDS_NO_NAME)
		for (int i1 = 1; i1 < i; i1++)
		{
			if (szdata == table1.Columns[i1])
			{
				CString s;
				s.Format(L"[%i]", i);
				szdata += s;
				break;
			}
		}
		pKeyCombo11->AddItem(szdata);
		pKeyCombo12->AddItem(szdata);
		pKeyCombo13->AddItem(szdata);
		table1.Columns[i] = szdata;
	}
}


void CChildView::OnKey11()
{
	table1.keys[0] = pKeyCombo11->GetCurSel() + 1;
}


void CChildView::OnKey12()
{
	table1.keys[1] = pKeyCombo12->GetCurSel() + 1;
}


void CChildView::OnKey13()
{
	table1.keys[2] = pKeyCombo13->GetCurSel() + 1;
}


void CChildView::OnPickSecondSheet()
{
	int tmpWSN = pSheetCombo2->GetCurSel() + 1;
	CString tmpWSS = pSheetCombo2->GetEditText();
	g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)); // CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)
	if (tmpWSN > 0)
	{
		saRet2.Destroy();
		table2.WorkSheetNumber = tmpWSN;
		sheet2 = sheets2.get_Item(COleVariant(tmpWSS));
		oRange2 = sheet2.get_UsedRange();

		saRet2 = oRange2.get_Value(covOptional);
		long iRows;
		long iCols;
		saRet2.GetUBound(1, &iRows);
		saRet2.GetUBound(2, &iCols);
		table2.NumberOfColumns = iCols;
		table2.NumberOfRows = iRows;
		table2.RowWithNames = 1;

		CString tmps;
		tmps.Format(_T("%d"), 1);
		pSpinner2_Names->SetEditText(tmps);
		table2.RowWithNames = 1;

		tmps.Format(_T("%d"), 2);
		pSpinner2_Fdata->SetEditText(tmps);
		table2.FirstRowWithData = 2;

		m_nCellWidth = STEP_X;
		m_nCellHeight = STEP_Y;
		m_nRibbonWidth = 0;
		m_nViewWidth = STEP_X + OFFSET_X + ((table2.NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
		m_nViewHeight = STEP_Y + OFFSET_Y + m_nCellHeight * (table1.NumberOfColumns + 1);

		SCROLLINFO si;
		si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
		si.nMin = 0;
		si.nMax = m_nViewWidth - 1;
		si.nPos = m_nHScrollPos;
		si.nPage = m_nHPageSize;

		SetScrollInfo(SB_HORZ, &si, TRUE);

		table2.keys[0] = 0;
		table2.keys[1] = 0;
		table2.keys[2] = 0;

		if (matrixDone > 0)
		{
			mxClear(table2.NumberOfColumns + 1, table1.NumberOfColumns + 1);
			matrixDone = 0;
		}

		updateCombos2();
		this->Invalidate();
		matrixDone = false;

		HWND hWnd0 = this->GetSafeHwnd();
		g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED)); // CMsg(IDS_DATA_VERIFIED)
		threadCnt++; AfxBeginThread(makePrereq2ThreadProc, hWnd0);
	}
}


void CChildView::updateCombos2()
{
	pKeyCombo21->RemoveAllItems();
	pKeyCombo22->RemoveAllItems();
	pKeyCombo23->RemoveAllItems();
	pKeyCombo21->SetEditText(_T(""));
	pKeyCombo22->SetEditText(_T(""));
	pKeyCombo23->SetEditText(_T(""));
	long index[2];
	CString szdata;
	COleVariant vData;
	for (int i = 1; i <= table2.NumberOfColumns; i++)
	{
		index[0] = table2.RowWithNames;
		index[1] = i;
		try {
			saRet2.GetElement(index, vData); vData = (CString)vData;
		}
		catch (COleException* e)
		{
			vData = L"";
		}
		szdata = vData;
		if (szdata == "") szdata = CMsg(IDS_NO_NAME); // CMsg(IDS_NO_NAME)
		for (int i1 = 1; i1 < i; i1++)
		{
			if (szdata == table2.Columns[i1])
			{
				CString s;
				s.Format(L"[%i]", i);
				szdata += s;
				break;
			}
		}
		pKeyCombo21->AddItem(szdata);
		pKeyCombo22->AddItem(szdata);
		pKeyCombo23->AddItem(szdata);
		table2.Columns[i] = szdata;
	}
}


void CChildView::OnUpdateSpin2Fdata(CCmdUI *pCmdUI)
{
	
	if (!(filename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSpinner2_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, pRibbon->FindByID(ID_SPIN2_FDATA));
}


void CChildView::OnSpin2Fdata()
{
	CString tmps = pSpinner2_Fdata->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 2) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	pSpinner2_Fdata->SetEditText(tmps);
	table2.FirstRowWithData = tmpi;
	prereq2valid = false;
}


void CChildView::OnUpdateSpin2Names(CCmdUI *pCmdUI)
{
	
	if (!(filename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSpinner2_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, pRibbon->FindByID(ID_SPIN2_NAMES));
}


void CChildView::OnSpin2Names()
{
	CString tmps = pSpinner2_Names->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 1) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	pSpinner2_Names->SetEditText(tmps);

	table2.RowWithNames = tmpi;

	updateCombos2();
	this->Invalidate();

}





void CChildView::OnUpdateKey21(CCmdUI *pCmdUI)
{
	
	if (!(filename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pKeyCombo21 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_KEY2_1));
}


void CChildView::OnUpdateKey22(CCmdUI *pCmdUI)
{
	
	if (!(filename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pKeyCombo22 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_KEY2_2));
}


void CChildView::OnUpdateKey23(CCmdUI *pCmdUI)
{
	
	if (!(filename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pKeyCombo23 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_KEY2_3));
}


void CChildView::OnKey21()
{
	table2.keys[0] = pKeyCombo21->GetCurSel() + 1;
}

void CChildView::OnKey22()
{
	table2.keys[1] = pKeyCombo22->GetCurSel() + 1;
}


void CChildView::OnKey23()
{
	table2.keys[2] = pKeyCombo23->GetCurSel() + 1;
}


void CChildView::makeCharArr1()
{
	if (int arSize1 = (table1.NumberOfColumns + 1) * (table1.NumberOfRows + 1))
	{
		long prgHlpr0, prgHlpr;
		prgHlpr0 = 0;
		prgHlpr = 0;

		delete[] mainArr1;
		mainArr1 = new char[arSize1];
		long index[2];
		char chr;
		COleVariant vData;
		CString szdata;
		for (int i_c = 1; i_c <= table1.NumberOfColumns; i_c++)
		{

			prgHlpr0 = 100 * i_c / table1.NumberOfColumns;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
			}

			for (int i_r = 1; i_r <= table1.NumberOfRows; i_r++)
			{
				index[0] = i_r;
				index[1] = i_c;
				try {
					saRet1.GetElement(index, vData); vData = (CString)vData;
					szdata = vData;
				}
				catch (COleException* e)
				{
					szdata = "";
				}
				if (szdata == "")
				{
					chr = 0;
				}
				else
				{
					chr = szdata[0];
				}
				mainArr1[(i_r - 1) * table1.NumberOfColumns + i_c] = chr;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 100);
}


void CChildView::makeCharArr2()
{
	if (int arSize2 = (table2.NumberOfColumns + 1) * (table2.NumberOfRows + 1))
	{
		long prgHlpr0, prgHlpr;
		prgHlpr0 = 0;
		prgHlpr = 0;

		delete[] mainArr2;
		mainArr2 = new char[arSize2];
		long index[2];
		char chr;
		COleVariant vData;
		CString szdata;
		for (int i_c = 1; i_c <= table2.NumberOfColumns; i_c++)
		{
			prgHlpr0 = 100 * i_c / table2.NumberOfColumns;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
			}
			//TRACE("i_c: %i\n", i_c);
			for (int i_r = 1; i_r <= table2.NumberOfRows; i_r++)
			{
				//TRACE("i_r: %i, i_c: %i\n", i_r, i_c);
				index[0] = i_r;
				index[1] = i_c;
				try {
					saRet2.GetElement(index, vData); vData = (CString)vData;
					szdata = vData;
				}
				catch (COleException* e)
				{
					szdata = "";
				}
				if (szdata == "")
				{
					chr = 0;
				}
				else
				{
					chr = szdata[0];
				}
				mainArr2[(i_r - 1) * table2.NumberOfColumns + i_c] = chr;
			}
		}

	}
	PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
}


void CChildView::OnLButtonDblClk(UINT nFlags, CPoint point)
{

	if (lockPrg1 || lockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}


	if (matrixDone && (cCell.y <= table1.NumberOfColumns && cCell.x <= table2.NumberOfColumns))
	{
		g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)

		lockPrg2 = true;
		HWND hWnd0 = this->GetSafeHwnd();

		
		int mx_X_max = table2.NumberOfColumns;
		markedMatrix[(cCell.y - 1) * mx_X_max + cCell.x] = true;
		this->Invalidate();
		int rslt;

		threadCnt++; AfxBeginThread(MyThreadProc3, hWnd0);
	}

	CWnd::OnLButtonDblClk(nFlags, point);
}


void CChildView::mxClear(int x, int y)
{
	int size = (x + 1) * (y + 1);
	delete[] mainMatrix;
	mainMatrix = new int[size];
	delete[] markedMatrix;
	markedMatrix = new bool[size];
	for (int i = 0; i < size; i++)
	{
		mainMatrix[i] = 0;
		markedMatrix[i] = false;
	}
}


int CChildView::mxPut(int x, int y)
{
	int mx_X_max = table2.NumberOfColumns;

	int index = (y - 1) * mx_X_max + x;

	mainMatrix[index] += 1;

	return 0;
}


int CChildView::mxGet(int x, int y)
{
	int mx_X_max = table2.NumberOfColumns;

	int index = (y - 1) * mx_X_max + x;

	return mainMatrix[index];
}

bool CChildView::mxMarkedGet(int x, int y)
{
	int mx_X_max = table2.NumberOfColumns;

	int index = (y - 1) * mx_X_max + x;

	return markedMatrix[index];
}


void CChildView::checkEmptiness1()
{
	delete[] emptyClms1;
	emptyClms1 = new bool[table1.NumberOfColumns + 2];

	for (int i = 0; i <= table1.NumberOfColumns; i++) emptyClms1[i] = true;

	long prgHlpr0, prgHlpr;
	prgHlpr0 = 0;
	prgHlpr = 0;


	for (int i_c = 1; i_c <= table1.NumberOfColumns; i_c++)
	{
		prgHlpr0 = 100 * i_c / table1.NumberOfColumns;
		if (prgHlpr0 > prgHlpr + 10)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
		}
		for (int i_r = table1.FirstRowWithData; i_r <= table1.NumberOfRows; i_r++)
		{
			if (mainArr1[(i_r - 1) * table1.NumberOfColumns + i_c])
			{
				emptyClms1[i_c] = false;
				break;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 100);
}

void CChildView::checkEmptiness2()
{
	delete[] emptyClms2;
	emptyClms2 = new bool[table2.NumberOfColumns + 2];

	for (int i = 0; i <= table2.NumberOfColumns; i++) emptyClms2[i] = true;

	long prgHlpr0, prgHlpr;
	prgHlpr0 = 0;
	prgHlpr = 0;


	for (int i_c = 1; i_c <= table2.NumberOfColumns; i_c++)
	{
		prgHlpr0 = 100 * i_c / table2.NumberOfColumns;
		if (prgHlpr0 > prgHlpr + 10)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
		}
		for (int i_r = table2.FirstRowWithData; i_r <= table2.NumberOfRows; i_r++)
		{
			if (mainArr2[(i_r - 1) * table2.NumberOfColumns + i_c])
			{
				emptyClms2[i_c] = false;
				break;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
}


bool CChildView::checkKeysUniqueness1()
{
	lockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	CString szTaken_A, szTaken_B;
	for (int i0 = table1.FirstRowWithData; i0 <= table1.NumberOfRows; i0++)
	{
		prgHlpr0 = 100 * i0 / table1.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
		}
		szTaken_A = keyArr11[i0];


		for (int i1 = i0 + 1; i1 <= table1.NumberOfRows; i1++)
		{
			szTaken_B = keyArr11[i1];


			if (szTaken_A == szTaken_B)
			{

				lockPrg1 = false;
				PostMessage(CM_UPDATE_PROGRESS, 0, 100);
				return false;
			}
		}
	}
	lockPrg1 = false;
	return true;
}


bool CChildView::checkKeysUniqueness2()
{
	lockPrg2 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	CString szTaken_A, szTaken_B;
	for (int i0 = table2.FirstRowWithData; i0 <= table2.NumberOfRows; i0++)
	{
		prgHlpr0 = 100 * i0 / table2.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
		}
		szTaken_A = keyArr21[i0];


		for (int i1 = i0 + 1; i1 <= table2.NumberOfRows; i1++)
		{
			szTaken_B = keyArr21[i1];


			if (szTaken_A == szTaken_B) 
			{

				lockPrg2 = false;
				PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
				return false;
			}
		}
	}
	lockPrg2 = false;
	return true;
}


void CChildView::firstPass()
{

	if (!prereq1valid) makePrereq1();
	if (!prereq2valid) makePrereq2();

	doAutoMark = autoMark;
	lockPrg1 = true;
	CString concatenatedKey1, concatenatedKey2;
	int prgHlpr = 0, prgHlpr0 = 0;
	char firstChar1, firstChar2;
	effMax = 0;

	mxClear(table2.NumberOfColumns + 1, table1.NumberOfColumns + 1);

	POSITION mapPos1;

	mapPos1 = map1.GetStartPosition();

	//// The commented code below is used only if the keys are stored in arrays instead of maps

	long /*keyRow1, */ keyRow2;

	int fchar1_y, fchar2_y;

	//int i1; // iterator for progress visualisation;
	//i1 = table1.FirstRowWithData-1;
	if (autoMark)
	{
		for (long i1 = table1.FirstRowWithData; i1 <= table1.NumberOfRows; i1++)
			//while (mapPos1 !=  NULL)
		{
			//i1++;
			prgHlpr0 = 99 * i1 / table1.NumberOfRows; // 99: because 100 would terminate the thread immaturely
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				//pProgressBar1->SetPos(prgHlpr);
			}

			//map1.GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);

			concatenatedKey1 = keyArr11[i1];
			//for (int i2 = table2.FirstRowWithData; i2 <= table2.NumberOfRows; i2++)
			//{

				//concatenatedKey2 = keyArr21[i2];

				//if (concatenatedKey1 == concatenatedKey2)
			if (map2.Lookup(concatenatedKey1, (long&)keyRow2))
			{
				effMax++;
				//procBoundaries[thrdIdx].effPortion++;
				fchar1_y = (i1 - 1) * table1.NumberOfColumns;
				for (int i3 = 1; i3 <= table1.NumberOfColumns; i3++)
				{
					firstChar1 = mainArr1[fchar1_y + i3];
					fchar2_y = (keyRow2 - 1) * table2.NumberOfColumns;
					for (int i4 = 1; i4 <= table2.NumberOfColumns; i4++)
					{
						firstChar2 = mainArr2[fchar2_y + i4];
						if (firstChar1 == firstChar2)
						{
							// empty combination OR ...
							//TRACE("i1 = %d and i3 = %d\n", i1, i3);

							if (firstChar1 == 0 || (getCellValue1(i3, i1) == getCellValue2(i4, keyRow2)))
							{
								mxPut(i4, i3);
							}
							else
							{

								if (table1.Columns[i3] == table2.Columns[i4])
								{
									mainArr1[fchar1_y + i3] = 1;
									mainArr2[fchar2_y + i4] = 1;
								}

							}

						}
						else
						{

							if (table1.Columns[i3] == table2.Columns[i4])
							{
								mainArr1[fchar1_y + i3] = 1;
								mainArr2[fchar2_y + i4] = 1;
							}

						}
					}
				}
			}
			else
			{
				keyMissing1[i1] = true;
			}
			//}
		}
		if (in2file)
		{
			long keyRow1;
			prgHlpr = 0; prgHlpr0 = 0;
			for (long i1_2 = table2.FirstRowWithData; i1_2 <= table2.NumberOfRows; i1_2++)
			{
				prgHlpr0 = 100 * i1_2 / table2.NumberOfRows;
				if (prgHlpr0 > prgHlpr)
				{
					prgHlpr = prgHlpr0;
					PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				}
				concatenatedKey2 = keyArr21[i1_2];
				if (!map1.Lookup(concatenatedKey2, (long&)keyRow1))
				{
					keyMissing2[i1_2] = true;
				}
			}
		}
		PostMessage(CM_UPDATE_PROGRESS, 0, 100); // because otherwise the "resolve auto mark" procedure would be started prematurely
	}
	else
	{
		for (long i1 = table1.FirstRowWithData; i1 <= table1.NumberOfRows; i1++)
			//while (mapPos1 !=  NULL)
		{
			//i1++;
			prgHlpr0 = 100 * i1 / table1.NumberOfRows;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				//pProgressBar1->SetPos(prgHlpr);
			}

			//map1.GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);

			concatenatedKey1 = keyArr11[i1];
			//for (int i2 = table2.FirstRowWithData; i2 <= table2.NumberOfRows; i2++)
			//{

			//concatenatedKey2 = keyArr21[i2];

			//if (concatenatedKey1 == concatenatedKey2)
			if (map2.Lookup(concatenatedKey1, (long&)keyRow2))
			{
				effMax++;
				//procBoundaries[thrdIdx].effPortion++;
				fchar1_y = (i1 - 1) * table1.NumberOfColumns;
				for (int i3 = 1; i3 <= table1.NumberOfColumns; i3++)
				{
					firstChar1 = mainArr1[fchar1_y + i3];
					fchar2_y = (keyRow2 - 1) * table2.NumberOfColumns;
					for (int i4 = 1; i4 <= table2.NumberOfColumns; i4++)
					{
						firstChar2 = mainArr2[fchar2_y + i4];
						if (firstChar1 == firstChar2)
						{
							// empty combination OR ...
							//TRACE("i1 = %d and i3 = %d\n", i1, i3);

							if (firstChar1 == 0 || (getCellValue1(i3, i1) == getCellValue2(i4, keyRow2)))
							{
								mxPut(i4, i3);
							}
						}
					}
				}
			}
			//}
		}
	}
	VARIANT val;

	delete[] greenClms1;
	greenClms1 = new bool[table1.NumberOfColumns + 2];
	delete[] greenClms2;
	greenClms2 = new bool[table2.NumberOfColumns + 2];

	for (int i = 0; i <= table1.NumberOfColumns; i++) greenClms1[i] = false;
	for (int i = 0; i <= table2.NumberOfColumns; i++) greenClms2[i] = false;

	for (int i_c = 1; i_c <= table2.NumberOfColumns; i_c++)
	{
		for (int i_r = 1; i_r <= table1.NumberOfColumns; i_r++)
		{
			if (mxGet(i_c, i_r) == effMax)
			{
				greenClms1[i_r] = true;
				greenClms2[i_c] = true;
			}
		}
	}

	matrixDone++;
	PostMessage(CM_UPDATE_PROGRESS, 0, 1000);




	lockPrg1 = false;
}


int CChildView::createKeyArrays1()
{
	notUniqueKeys = { 0, 0, L"" };

	long mapIdx;
	long index[2];
	char chr;
	COleVariant vData;
	CString szdata;

	map1.RemoveAll();

	delete[] keyArr11;
	keyArr11 = new CString[table1.NumberOfRows + 2];
	delete[] keyMissing1;
	keyMissing1 = new bool[table1.NumberOfRows + 2];

	lockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;

	for (int i_i = table1.FirstRowWithData; i_i <= table1.NumberOfRows; i_i++)
	{
		prgHlpr0 = 100 * i_i / table1.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
		}

		szdata = "";
		if (table1.keys[0])
		{

			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = table1.keys[0];
			try {
				saRet1.GetElement(index, vData); vData = (CString)vData;
			} 
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata+= vData;
		}
		if (table1.keys[1])
		{

			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = table1.keys[1];
			try {
				saRet1.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata+= vData;
		}
		if (table1.keys[2])
		{

			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = table1.keys[2];
			try {
				saRet1.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata+= vData;
		}
		keyArr11[i_i] = szdata;
		keyMissing1[i_i] = false;
		if (map1.Lookup(szdata, (long&)mapIdx))
		{
			notUniqueKeys = { i_i, mapIdx, szdata };
			map1.RemoveAll();
			return 1;
		}
		map1.SetAt(szdata, i_i);

	}

	PostMessage(CM_UPDATE_PROGRESS, 0, 1000);
	return 0;
}

int CChildView::createKeyArrays2()
{
	notUniqueKeys = { 0, 0, L"" };

	long mapIdx;
	long index[2];
	char chr;
	COleVariant vData;
	CString szdata;

	map2.RemoveAll();

	delete[] keyArr21;
	keyArr21 = new CString[table2.NumberOfRows + 2];
	delete[] keyMissing2;
	keyMissing2 = new bool[table2.NumberOfRows + 2];

	lockPrg2 = true;
	int prgHlpr = 0, prgHlpr0 = 0;

	prgHlpr = 0;
	prgHlpr0 = 0;

	for (int i_i = table2.FirstRowWithData; i_i <= table2.NumberOfRows; i_i++)
	{
		prgHlpr0 = 100 * i_i / table2.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
		}

		szdata = "";
		if (table2.keys[0])
		{

			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = table2.keys[0];
			try {
				saRet2.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata += vData;
		}
		if (table2.keys[1])
		{
			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = table2.keys[1];
			try {
				saRet2.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata += vData;
		}
		if (table2.keys[2])
		{

			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = table2.keys[2];
			try {
				saRet2.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata += vData;
		}
		keyArr21[i_i] = szdata;
		keyMissing2[i_i] = false;
		if (map2.Lookup(szdata, (long&)mapIdx))
		{
			notUniqueKeys = { i_i, mapIdx, szdata };
			map2.RemoveAll();
			return 2;
		}
		map2.SetAt(szdata, i_i);
	}
	PostMessage(CM_UPDATE_PROGRESS2, 0, 1000);
	return 0;
}

CString CChildView::getCellValue1(int column, int row)
{

	long index[2];
	COleVariant vData;
	CString szdata;

	index[0] = row;
	index[1] = column;
	try {
		saRet1.GetElement(index, vData); vData = (CString)vData;
	}
	catch (COleException* e)
	{
		vData = L"";
	}
	szdata = vData;

	return szdata;
}


CString CChildView::getCellValue2(int column, int row)
{

	long index[2];
	COleVariant vData;
	CString szdata;

	index[0] = row;
	index[1] = column;
	try {
		saRet2.GetElement(index, vData); vData = (CString)vData;
	}
	catch (COleException* e)
	{
		vData = L"";
	}
	szdata = vData;

	return szdata;
}


void CChildView::OnMouseMove(UINT nFlags, CPoint point)
{
	ChosenCell oldCell;
	oldCell.x = cCell.x;
	oldCell.y = cCell.y;

	if (point.x > OFFSET_X + STEP_X)
	{
		cCell.x = (point.x - OFFSET_X) / STEP_X + visTopLeft.left; 
	}
	else
	{
		cCell.x = 0;
	}
	if (point.y > OFFSET_Y + STEP_Y)
	{
		cCell.y = (point.y - OFFSET_Y) / STEP_Y + visTopLeft.top;
	}
	else
	{
		cCell.y = 0;
	}

	CString s;
	CString sx, sy;
	sx.Format(L"%i", cCell.x);
	sy.Format(L"%i", cCell.y);
	sx = CMsg(IDS_COORDS);
	s.Format(CMsg(IDS_COORDS), cCell.y, cCell.x, threadCnt); // CMsg(IDS_COORDS)

	g_pMainFrame->updateStatusBar(s);

	if (!(oldCell.x == cCell.x) || !(oldCell.y == cCell.y))
	{
		if (!forceNotOnlyPcnt)
		{
			onlyPcnt = true;
		}
		else
		{
			onlyPcnt = false;
			forceNotOnlyPcnt = false;
		}
		RECT rct;
		rct.left = 0; rct.top = 0; rct.right = OFFSET_X + STEP_X - 2; rct.bottom = OFFSET_Y + STEP_Y - 2;
		this->InvalidateRect(&rct, 1);
	}
	this->SetFocus();
}


void CChildView::OnSlider2()
{
	sldr = pSlider->GetPos();
	this->Invalidate();

	CString s;
	CString sx;
	s = rsltTxt;
	sx.Format(CMsg(IDS_MARK_SUSP_INTERS), pSlider->GetPos()); // CMsg(IDS_MARK_SUSP_INTERS)
	s = sx + L" %";

	g_pMainFrame->updateStatusBar(s);
}


void CChildView::OnUpdateSlider2(CCmdUI *pCmdUI)
{
	
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSlider = DYNAMIC_DOWNCAST(CMFCRibbonSlider, pRibbon->FindByID(ID_SLIDER2));
	if (pSlider->GetPos() == 0)
		pSlider->SetPos(sldr);
}


void CChildView::OnCheck4()
{
	in1file = !in1file;
}


void CChildView::OnUpdateCheck4(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(in1file);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pMarkIn1 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, pRibbon->FindByID(ID_CHECK4));
}


void CChildView::OnCheck5()
{
	in2file = !in2file;
}


void CChildView::OnUpdateCheck5(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(in2file);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pMarkIn2 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, pRibbon->FindByID(ID_CHECK5));
}


void CChildView::OnButton2()
{
	if (app)
	{
		app.put_Visible(TRUE);
		app.put_UserControl(TRUE);
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
	CRange range = sheet1.get_Range(COleVariant(cnv), COleVariant(cnv));
	interior = range.get_Interior();
	interior.put_Color(COleVariant(long(RGB(palette[chosenColor1].red, palette[chosenColor1].green, palette[chosenColor1].blue))));
	return;
}


void CChildView::markIn2(int row, int clm)
{
	CString cnv = convertR1C1(row, clm);
	CRange range = sheet2.get_Range(COleVariant(cnv), COleVariant(cnv));
	interior = range.get_Interior();
	interior.put_Color(COleVariant(long(RGB(palette[chosenColor2].red, palette[chosenColor2].green, palette[chosenColor2].blue))));
	return;
}



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
		visTopLeft.left = 0;
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
		visTopLeft.top = 0;
	}

	si.fMask = SIF_PAGE | SIF_RANGE | SIF_POS;
	si.nMin = 0;
	si.nMax = nVScrollMax;
	si.nPos = m_nVScrollPos;
	si.nPage = m_nVPageSize;

	SetScrollInfo(SB_VERT, &si, TRUE);
	onlyPcnt = false;
	//this->Invalidate(); // uncomment in case of problems with redrawing after RESIZE

	visTopLeft.top = m_nVScrollPos / STEP_Y;
	SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
	RECT rect;
	GetClientRect(&rect);

	rect.top = OFFSET_Y + STEP_Y;
	ScrollWindow(0, 0, &rect);
	onlyPcnt = false;
	forceNotOnlyPcnt = true;
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
	m_nViewWidth = STEP_X + OFFSET_X + ((table2.NumberOfColumns + 1) * m_nCellWidth) + m_nRibbonWidth;
	m_nViewHeight = STEP_Y + OFFSET_Y  + m_nCellHeight * (table1.NumberOfColumns + 1);
	sldr = 90;


	return 0;
}


void CChildView::OnUpdateProgress2(CCmdUI *pCmdUI)
{
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pProgressBar2 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, pRibbon->FindByID(ID_PROGRESS3));
	// Emergency update of the container for found differences
	pFoundDifferences = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, pRibbon->FindByID(ID_DIFFS_LIST));
	pToFront = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, pRibbon->FindByID(ID_PUT_TO_FRONT));
}


void CChildView::OnUpdateCheck2(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(verifyKeys);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pVerifyKeys = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, pRibbon->FindByID(ID_CHECK2));
}

void CChildView::OnCheck2()
{
	verifyKeys = !verifyKeys;
}

void CChildView::OnUpdateButton2(CCmdUI *pCmdUI)
{
	if (!&app) pCmdUI->Enable(false); else pCmdUI->Enable(true);
}

void CChildView::OnCheck7()
{
	sameNames = !sameNames;

		visTopLeft.top = m_nVScrollPos / STEP_Y;
		SetScrollPos(SB_VERT, m_nVScrollPos, TRUE);
		RECT rect;
		GetClientRect(&rect);

		rect.top = OFFSET_Y + STEP_Y;
		ScrollWindow(0, 0, &rect);
		onlyPcnt = false;
		forceNotOnlyPcnt = true;
		this->Invalidate();
}


void CChildView::OnUpdateCheck7(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(sameNames);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pSameNames = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, pRibbon->FindByID(ID_CHECK7));
}

UINT MyThreadProc(LPVOID pParam)
{

	HWND hWnd1 = (HWND)pParam;

	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);

	pWnd->firstPass();
	

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}

afx_msg LRESULT CChildView::OnCmUpdateProgress(WPARAM wParam, LPARAM lParam)
{
	if ((UINT)lParam > 99)
	{
			pProgressBar1->SetPos(0);
			lockPrg1 = false;
			this->Invalidate();
			if (doAutoMark)
			{
				g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_IN_EXCEL_RUNNING)); // CMsg(IDS_MARKING_IN_EXCEL_RUNNING)
				resolveAutoMark();
				g_pMainFrame->updateStatusBar(CMsg(IDS_DONE)); // CMsg(IDS_DONE)
			}
			if ((UINT)lParam == 1000)
			{
				if (waitingForKeys)
				{
					keys1done = true;
					if (keys2done)
					{
						HWND hWnd0 = this->GetSafeHwnd();
						waitingForKeys = false;
						keys1done = false;
						keys2done = false;
						threadCnt++; AfxBeginThread(MyThreadProc, hWnd0);
						g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
					}
				}
				else
				{
					rsltTxt.Format(CMsg(IDS_FOUND_KEYS_FROM_TOTAL), effMax, (table1.NumberOfRows - table1.FirstRowWithData + 1), (table2.NumberOfRows - table2.FirstRowWithData + 1)); // CMsg(IDS_FOUND_KEYS_FROM_TOTAL)
					g_pMainFrame->updateStatusBar(rsltTxt);
				}
			}
	}
	else
	{
		pProgressBar1->SetPos((UINT)lParam);
	}
	return 0;
}


afx_msg LRESULT CChildView::OnCmUpdateProgress2(WPARAM wParam, LPARAM lParam)
{
	if ((UINT)lParam > 99)
	{
		pProgressBar2->SetPos(0);
		lockPrg2 = false;
		this->Invalidate();

		if ((UINT)lParam == 1000)
		{
			if (waitingForKeys)
			{
				keys2done = true;
				if (keys1done)
				{
					HWND hWnd0 = this->GetSafeHwnd();
					waitingForKeys = false;
					keys1done = false;
					keys2done = false;
					threadCnt++; AfxBeginThread(MyThreadProc, hWnd0);
					g_pMainFrame->updateStatusBar(CMsg(IDS_X_COMP_IN_PRGRS)); // CMsg(IDS_X_COMP_IN_PRGRS)
				}
			}
		}
	}
	else
	{
		pProgressBar2->SetPos((UINT)lParam);
	}
	return 0;
}


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
		g_pMainFrame->updateStatusBar(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING));  // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		pProgressBar1->SetPos(0);
		pFoundDifferences->RemoveAllItems();
		pFoundDifferences->SetEditText(L"");

		{
			nor = table1.NumberOfRows + 1;
			for (int i1 = 1; i1 < nor; i1++)
			{
				prgHlpr = (i1 * 100) / nor;
				if (prgHlpr > prgHlpr0)
				{
					SendMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
					prgHlpr0 = prgHlpr;
				}
				dfrncRow2 = foundDifferences[i1];
				if (dfrncRow2 > 0)
				{
					if (++dfrnCntr < 500)
					{
						fndDfrnc1 = L"";
						fndDfrnc1.Format(L"(1r%i):", i1);
						fndDfrnc1 += getCellValue1(oldy, i1);
						fndDfrnc1 = fndDfrnc1.Left(26);
						fndDfrnc2 = L"";
						fndDfrnc2.Format(L"   (2r%i):", dfrncRow2);
						fndDfrnc2 += getCellValue2(oldx, dfrncRow2);
						fndDfrnc2 = fndDfrnc2.Left(26);
						selKey = L"";
						selKey.Format(L"%s%s   (key): %s", fndDfrnc1, fndDfrnc2, keyArr11[i1]);
						fndDfrnc = selKey.Left(54);

						//fndDfrnc = fndDfrnc1 + fndDfrnc2 + selKey;

						
						pFoundDifferences->AddItem((LPCTSTR)fndDfrnc);
					}
				}
				if (in1file)
				{
					if (markIn1Arr[i1])
					{
	
						if (starts == L"")
						{
							starts = convertR1C1(i1, oldy);
						}
						ends = convertR1C1(i1, oldy);

					}
					else
					{
						if (!(starts == L"") && !(ends == L""))
						{
							CRange range = sheet1.get_Range(COleVariant(starts), COleVariant(ends));
							interior = range.get_Interior();
							interior.put_Color(COleVariant(long(RGB(palette[chosenColor1].red, palette[chosenColor1].green, palette[chosenColor1].blue))));
							starts = L"";
							ends = L"";
						}
					}
				}
			}
			if (in1file && !(starts == L"") && !(ends == L""))
			{
				CRange range = sheet1.get_Range(COleVariant(starts), COleVariant(ends));
				interior = range.get_Interior();
				interior.put_Color(COleVariant(long(RGB(palette[chosenColor1].red, palette[chosenColor1].green, palette[chosenColor1].blue))));
				starts = L"";
				ends = L"";
			}
		}
		temps = L"";
		starts = L"";
		ends = L"";
		if (in2file)
		{
			nor = table2.NumberOfRows + 1;
			for (int i2 = 1; i2 < nor; i2++)
			{
				prgHlpr = (i2 * 100) / nor;
				if (prgHlpr > prgHlpr0)
				{
					SendMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
					prgHlpr0 = prgHlpr;
				}
				if (markIn2Arr[i2])
				{

					if (starts == L"")
					{
						starts = convertR1C1(i2, oldx);
					}
					ends = convertR1C1(i2, oldx);

				}
				else
				{
					if (!(starts == L"") && !(ends == L""))
					{
						CRange range = sheet2.get_Range(COleVariant(starts), COleVariant(ends));
						interior = range.get_Interior();
						interior.put_Color(COleVariant(long(RGB(palette[chosenColor2].red, palette[chosenColor2].green, palette[chosenColor2].blue))));
						starts = L"";
						ends = L"";
					}
				}
			}
			if (!(starts == L"") && !(ends == L""))
			{
				CRange range = sheet2.get_Range(COleVariant(starts), COleVariant(ends));
				interior = range.get_Interior();
				interior.put_Color(COleVariant(long(RGB(palette[chosenColor2].red, palette[chosenColor2].green, palette[chosenColor2].blue))));
				starts = L"";
				ends = L"";
			}
		}

		SendMessage(CM_UPDATE_PROGRESS2, 0, 100);
		lockPrg2 = false;
		g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE)); // CMsg(IDS_MARKING_DONE)
		EndWaitCursor();
		DrainMsgQueue();
	}
	else
	{
		pProgressBar1->SetPos((UINT)lParam);
	}
	
	return 0;
}

UINT MyThreadProc2(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	uniqueKeys1 = false;
	uniqueKeys2 = false;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;
	rslt = pWnd->createKeyArrays1();
	if (rslt == 1)
	{
		pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE)); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE)
		return 0;
	}
	uniqueKeys1 = true;
	rslt = pWnd->createKeyArrays2();
	if (rslt == 2)
	{
		pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE));// CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE)
		return 0;
	}

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}

UINT CreateKeys1ThreadProc(LPVOID pParam)
{

	HWND hWnd1 = (HWND)pParam;
	uniqueKeys1 = false;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;
	rslt = pWnd->createKeyArrays1();
	if (rslt == 1)
	{
		CString s;
		NotUniqueKeys* nu;
		nu = &notUniqueKeys;
		s.Format(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow, nu->secondRow); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS)
		pWnd->MessageBox(s);
		lockPrg1 = false;
		return 0;
	}

	uniqueKeys1 = true;

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}

UINT CreateKeys2ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	uniqueKeys2 = false;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;
	rslt = pWnd->createKeyArrays2();
	if (rslt == 2)
	{
		CString s;
		NotUniqueKeys* nu;
		nu = &notUniqueKeys;
		s.Format(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow, nu->secondRow); // CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS)
		pWnd->MessageBox(s);
		lockPrg2 = false;
		return 0;

	}

	uniqueKeys2 = true;

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}

UINT makePrereq1ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->makePrereq1();

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}

UINT makePrereq2ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->makePrereq2();

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}

UINT MyThreadProc3(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;

	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;

	pWnd->markInFiles();

	pWnd->decrementThreadCnt();
	AfxEndThread(0);
	return 0;

}



void CChildView::markInFiles()
{
		
		lockPrg2 = true;

		CString concatenatedKey1, concatenatedKey2;
		int prgHlpr = 0, prgHlpr0 = 0;
		char firstChar1, firstChar2;

		int cx, cy;
		cx = cCell.x;
		cy = cCell.y;
		oldy = cy;
		oldx = cx;

		delete[] markIn1Arr;
		markIn1Arr = new bool[table1.NumberOfRows + 2];
		delete[] markIn2Arr;
		markIn2Arr = new bool[table2.NumberOfRows + 2];
		delete[] foundDifferences;
		foundDifferences = new long[table1.NumberOfRows + 2];

		for (int i1 = 0; i1 <= table1.NumberOfRows + 1; i1++)
		{
			markIn1Arr[i1] = false;
			foundDifferences[i1] = 0;
		}

		for (int i2 = 0; i2 <= table2.NumberOfRows + 1; i2++)
		{
			markIn2Arr[i2] = false;
		}

		long keyRow1, keyRow2;

		POSITION mapPos1;
		mapPos1 = map1.GetStartPosition();

		int i1; // iterator for progress visualisation;
		i1 = table1.FirstRowWithData - 1;
		
		while (mapPos1 != NULL)
		{
			i1++;
			prgHlpr0 = 100 * i1 / table1.NumberOfRows;
			if (prgHlpr0 > prgHlpr) 
			{
				prgHlpr = prgHlpr0;

				PostMessage(CM_UPDATE_PROGRESS3, 0, prgHlpr);

			}
			map1.GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);

				if (map2.Lookup(concatenatedKey1, (long&)keyRow2))
				{

					if (!(getCellValue1(cy, keyRow1) == getCellValue2(cx, keyRow2)))
					{
						foundDifferences[keyRow1] = keyRow2;
						if (in1file) markIn1Arr[keyRow1] = true; //markIn1(i1, cy);
						if (in2file) markIn2Arr[keyRow2] = true; //markIn2(i2, cx);

					}
				}
		}
		PostMessage(CM_UPDATE_PROGRESS3, 0, 1000);
		lockPrg2 = false;

}

void CChildView::OnButton5()
{
	COLORREF i = (int)pColorPicker1->GetSelectedItem();
	chosenColor1 = i;
}


void CChildView::OnUpdateButton5(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(true);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pColorPicker1 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, pRibbon->FindByID(ID_BUTTON5));
}


void CChildView::OnButton3()
{
	COLORREF i = (int)pColorPicker2->GetSelectedItem();
	chosenColor2 = i;
}


void CChildView::OnUpdateButton3(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(true);
	pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	pColorPicker2 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, pRibbon->FindByID(ID_BUTTON3));
}


void CChildView::OnCheck3()
{
	autoMark = !autoMark;
}


void CChildView::OnUpdateCheck3(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(true);
	pCmdUI->SetCheck(autoMark);
}


void CChildView::makePrereq1()
{
	prereq1valid = false;
	delete[] mainArr1;
	mainArr1 = new char[table1.NumberOfRows + 2];

	makeCharArr1();

	checkEmptiness1();

	prereq1valid = true;
}


void CChildView::makePrereq2()
{
	prereq2valid = false;
	delete[] mainArr2;
	mainArr2 = new char[table2.NumberOfRows + 2];

	makeCharArr2();

	checkEmptiness2();

	prereq2valid = true;
}


void CChildView::resolveAutoMark()
{
	doAutoMark = false;
	g_pMainFrame->updateStatusBar(CMsg(IDS_DURING_MARKING_THREAD_BLOCKED));  // CMsg(IDS_DURING_MARKING_THREAD_BLOCKED)
	lockPrg2 = true;
	HWND hWnd = this->GetSafeHwnd();
	long nor, noc;
	int prgHlpr_x, prgHlpr0_x, prgHlpr_y, prgHlpr0_y;
	prgHlpr_x = 0;
	prgHlpr0_x = 0;
	prgHlpr_y = 0;
	prgHlpr0_y = 0;
	CString starts = L"";
	CString ends = L"";

		pProgressBar1->SetPos(0);

		BeginWaitCursor();
		for (int c1 = 1; c1 <= table1.NumberOfColumns; c1++)
		{
			prgHlpr0_x = 90 * c1 / table1.NumberOfColumns;
			if (prgHlpr0_x > prgHlpr_x)
			{
				prgHlpr_x = prgHlpr0_x;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr_x);
			}
			for (int c2 = 1; c2 <= table2.NumberOfColumns; c2++) 
			{
				if (table1.Columns[c1] == table2.Columns[c2])
				{
					if (in1file)
					{
						prgHlpr_y = 0;
						prgHlpr0_y = 0;
						for (long r1 = table1.FirstRowWithData; r1 <= table1.NumberOfRows; r1++)
						{
							prgHlpr0_y = 100 * r1 / table1.NumberOfRows;
							if (prgHlpr0_y > prgHlpr_y+10)
							{
								prgHlpr_y = prgHlpr0_y;
								PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr_y);
							}
							if (mainArr1[(r1 - 1) * table1.NumberOfColumns + c1] == 1)
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
									CRange range = sheet1.get_Range(COleVariant(starts), COleVariant(ends));
									interior = range.get_Interior();
									interior.put_Color(COleVariant(long(RGB(palette[chosenColor1].red, palette[chosenColor1].green, palette[chosenColor1].blue))));
									starts = L"";
									ends = L"";
								}
							}

						}
						if (!(starts == L"") && !(ends == L""))
						{
							CRange range = sheet1.get_Range(COleVariant(starts), COleVariant(ends));
							interior = range.get_Interior();
							interior.put_Color(COleVariant(long(RGB(palette[chosenColor1].red, palette[chosenColor1].green, palette[chosenColor1].blue))));
							starts = L"";
							ends = L"";
						}
					}
					starts = L"";
					ends = L"";
					if (in2file)
					{
						prgHlpr_y = 0;
						prgHlpr0_y = 0;
						for (long r2 = table2.FirstRowWithData; r2 <= table2.NumberOfRows; r2++)
						{
							prgHlpr0_y = 100 * r2 / table2.NumberOfRows;
							if (prgHlpr0_y > prgHlpr_y+10)
							{
								prgHlpr_y = prgHlpr0_y;
								PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr_y);
							}
							if (mainArr2[(r2 - 1) * table2.NumberOfColumns + c2] == 1)
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
									CRange range = sheet2.get_Range(COleVariant(starts), COleVariant(ends));
									interior = range.get_Interior();
									interior.put_Color(COleVariant(long(RGB(palette[chosenColor2].red, palette[chosenColor2].green, palette[chosenColor2].blue))));
									starts = L"";
									ends = L"";
								}
							}
						}
						if (!(starts == L"") && !(ends == L""))
						{
							CRange range = sheet2.get_Range(COleVariant(starts), COleVariant(ends));
							interior = range.get_Interior();
							interior.put_Color(COleVariant(long(RGB(palette[chosenColor2].red, palette[chosenColor2].green, palette[chosenColor2].blue))));
							starts = L"";
							ends = L"";
						}
					}
				}
			}
		}


		for (long r1 = table1.FirstRowWithData; r1 <= table1.NumberOfRows; r1++)
		{
			if (keyMissing1[r1])
			{
				starts = convertR1C1(r1, 1);
				ends = convertR1C1(r1, table1.NumberOfColumns);
				CRange range = sheet1.get_Range(COleVariant(starts), COleVariant(ends));
				interior = range.get_Interior();
				interior.put_Color(COleVariant(long(RGB(palette[chosenColor1].red, palette[chosenColor1].green, palette[chosenColor1].blue))));
			}
		}
		for (long r2 = table2.FirstRowWithData; r2 <= table2.NumberOfRows; r2++)
		{
			if (keyMissing2[r2]) // c1 - because we need it to run just once
			{
				starts = convertR1C1(r2, 1);
				ends = convertR1C1(r2, table2.NumberOfColumns);
				CRange range = sheet2.get_Range(COleVariant(starts), COleVariant(ends));
				interior = range.get_Interior();
				interior.put_Color(COleVariant(long(RGB(palette[chosenColor2].red, palette[chosenColor2].green, palette[chosenColor2].blue))));
			}
		}

		PostMessage(CM_UPDATE_PROGRESS, 0, 100);
		PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
		lockPrg2 = false;
		EndWaitCursor();
		g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE)); // CMsg(IDS_MARKING_DONE)
		DrainMsgQueue();
}

void CChildView::DrainMsgQueue(void)
{

	MSG     msg = { 0 };
	HWND hWnd = this->GetSafeHwnd();
	while (PeekMessage(&msg, hWnd, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE));

}



void CChildView::OnDiffslist()
{
	// there is no required answer for this event - at least now
}


void CChildView::OnUpdateDiffslist(CCmdUI *pCmdUI)
{
	// there is no required answer for this event - at least now
}

void CChildView::OnSel1()
{
	long row;
	long column;
	row = rowFromCombo();
	if (row > 0)
	{
		column = oldy;
		CString cnv = convertR1C1(row, column);
		CRange range = sheet1.get_Range(COleVariant(cnv), COleVariant(cnv));
		sheet1.Activate();
		range.Select();
		if (toFront)
		{
			app.put_Interactive(true);
			HWND ehWnd = (HWND)app.get_Hwnd();
			::PostMessage(ehWnd, WM_SHOWWINDOW, SW_RESTORE, 0);
			::SetForegroundWindow(ehWnd);
		}
	}
}

int CChildView::rowFromCombo()
{
	if (pFoundDifferences->GetCurSel() > -1)
	{
		CString s;
		s = pFoundDifferences->GetEditText();
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

void CChildView::OnButton6()
{

		long row;
		long column;
		row = rowFromCombo();
		if (row > 0)
		{
			row = foundDifferences[row];
	
			column = oldx;
			CString cnv = convertR1C1(row, column);
			CRange range = sheet2.get_Range(COleVariant(cnv), COleVariant(cnv));
			sheet2.Activate();
			range.Select();
			if (toFront)
			{
				app.put_Interactive(true);
				HWND ehWnd = (HWND)app.get_Hwnd();
				::PostMessage(ehWnd, WM_SHOWWINDOW, SW_RESTORE, 0);
				::SetForegroundWindow(ehWnd);
			}
		}
}

void CChildView::OnPut2front()
{
	toFront = !toFront;
}

void CChildView::OnUpdatePut2front(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(toFront);
}


void CChildView::decrementThreadCnt()
{
	threadCnt--;
}
