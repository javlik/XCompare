#include "stdafx.h"
#include <cstring>
#include <map>
#include <vector>
#include "XCompare.h"
#include "ChildView.h"
#include "MainFrm.h"
#include "Msg.h"
extern CMainFrame* g_pMainFrame; // pointer to FrameWindow
#define LINESIZE 8 // thickness of thick lines drawn in the visual representation of the result matrix
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#define swap(a,b) (a ^= b), (b ^= a), (a ^= b);
#define sgn(x) ( (int) ( (x > 0) - (x < 0) ) )
#define SUGKEYS 10
#define MAX_ATTEMPTS 1000000
#define MAX_CFileDialog_FILE_COUNT 1 // Number of selectable files in file selectors
#define FILE_LIST_BUFFER_SIZE ((MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1) 
#define CM_UPDATE_PROGRESS WM_APP + 1
#define CM_UPDATE_PROGRESS2 WM_APP + 2
#define CM_UPDATE_PROGRESS3 WM_APP + 3
#define CM_UPDATE_KEYPROGRESS1 WM_APP + 4
#define CM_UPDATE_KEYPROGRESS2 WM_APP + 5
#define OFFSET_Y 100 // height of a column (within the visual representation of the result matrix) that contains the names of column taken from the second table
#define OFFSET_X 100 // width of a row ... first table
#define STEP_X 24 // width of a cell ...
#define STEP_Y 24 // height of a cell ...
BOOL m_bUniqueKeys1; // to indicate whether values taken from key columns are identical - this is an inherent prerequisite for the comparison to be successful
BOOL m_bUniqueKeys2; // the same as above - for the second file
bool m_bWaitingForKeys; // indicates status of waiting for maps of keys
bool m_bKeys1done; // indicates status of readiness of keys for the first table
bool m_bKeys2done; // indicates status of readiness of keys for the second table
bool m_bKeysGathering1done;
bool m_bKeysGathering2done;
int m_nComplexity;
CString m_szRsltTxt; // human understandable text indicating some information related to the matrix - displayed in the status bar
struct PossibleKeys {
	int k[256];
};
//char autoKey1[256];
//char autoKey2[256];
unsigned long long m_nCheckedKeys1[MAX_ATTEMPTS + 1];
unsigned long long m_nCheckedKeys2[MAX_ATTEMPTS + 1];
int m_nCheckedKeysCounter1;
int m_nCheckedKeysCounter2;
struct Palette {
	int red;
	int green;
	int blue;
}; // structure type for color of cells
// int threadCnt;
struct Table {
	int WorkSheetNumber;
	long MaxNumberOfRows;
	long MaxNumberOfCols;	
	long NumberOfRows;
	int FirstRowWithData;
	int RowWithNames;
	int NumberOfColumns;
	CString Columns[256];
	bool keys[256];
	int keysCnt;
}; // structure type for description of tables
struct VisTopLeft {
	int top;
	int left;
}; // contains coords (in the units of matrix cells) of scrolled matrix
struct  KeyPair {
	int tab1;
	int tab2;
};
struct SimilaritiesAcrossTables {
	int clm1;
	int clm2;
	long similarity;
	int similarityOrder;
	int pureSim;
};
struct ChosenCell {
	int x;
	int y;
}; // contains coords (in the matrix cells) of the cell that is pointed by mouse
struct Clnt {
	int w;
	int h;
}; // Size of client area (in pixels)
struct BestKeyComb {
	int pk1;
	int pk2;
	int rating;
	long cnt;
};
struct NotUniqueKeys {
	long firstRow;
	long secondRow;
	CString keyString;
}; // if keys are found to not be unique, this structure contains the rows of the first found duplicate
NotUniqueKeys m_NotUniqueKeys1, m_NotUniqueKeys2;
//int threadCounter; // to be even more scallable in future
//int stepCounter;
Palette m_Palette[20]; // user can choose one of the 20 colors that will be used for background of found difference 
PossibleKeys m_PossibleKeys1[256]; // Contains combinations of found unique keys sorted by invEntropy
PossibleKeys m_PossibleKeys2[256];
KeyPair m_KeyPair[256]; // Unsorted pairs of keys
int		m_nKeyPairCounter; // The counter of key pairts - obviously
BestKeyComb m_BestKeyComb; // Found the most appropriate combination of keys
int m_nPossibleKeyCounter1 = 0; // Counter of possible keys - without respect to the other table
int m_nPossibleKeyCounter2 = 0; // 
long m_nInvEntropy1[256]; // Rating of "entropy" for found possible combinations - without respect to other table
long m_nInvEntropy2[256];
int m_nSortedEntropy1[256]; // Sorted rating of entropy - without respect to other table
int m_nSortedEntropy2[256];
bool m_bPrereq1valid, m_bPrereq2valid; // are prerequisities for execution of main process fulfilled
int m_nOldx, m_nOldy; // coords of last chosen cell
int m_nChosenColor1, m_nChosenColor2; // color will be used as a background in XLS file to mark difference
Clnt m_Clnt; // client area
bool m_bLockPrg1; // indicates status of computing in threads (inversely)
bool m_bLockPrg2; // ... same as above
bool m_bDoAutoMark; // whether user selected the option for automatic marking in XLS files
int m_nNatrixDone; // result matrix is done and ready for follow-up analysis
int m_nPrereqDone; // prerequisities for main comparison process are fulfilled
bool m_bMarkIdentCols;  // not used
bool m_bSameNames; // user wants the cells that intersects columns with same names to be marked by thick border
int m_nEffMax; // counted up number of keys that were found in both files
char *m_pchMainArr1; // this 2D array contains first character of content of each cell taken from the first XLS file (first horizontally, then vertically)
char *m_pchMainArr2; // .... the second XLS file
bool *m_pbMarkIn1Arr; // this 2D array indicates whether a cell at its coordinates (see above) is to be marked in the first file
bool *m_pbMarkIn2Arr; // ... in the second file.
CString *m_pszKeyArr11; // array of the strings found in the first key column in the first file
CString *m_pszKeyArr21; // ... the first key ... the second file
CString *m_pszTmpKeyArr11; // General temporary dynamic array of concatenated keys
CString *m_pszTmpKeyArr21;
bool *m_pbKeyMissing1; // General dynamic array of empty keys
bool *m_pbKeyMissing2; 
bool *m_pbTmpKeyMissing1; // General temporary dynamic array of empty keys
bool *m_pbTmpKeyMissing2;
int m_nExaminedKeys1[SUGKEYS + 4]; // Array of keys that were (or were not yet) checked for their inv entropy
int m_nExaminedKeys2[SUGKEYS + 4];
int m_nTmpKeys1[SUGKEYS + 4]; // Temporary array of keys that were checked for entropy - this workaround protects against possible collision of threads (in cpu cache)
int m_nTmpKeys2[SUGKEYS + 4];
int *m_pnMainMatrix; // 2D array representing the result matrix
bool *m_pbMarkedMatrix; // 2D array indicating marked cells in the result matrix
bool *m_pbEmptyClms1; // 1D array indicating empty columns in the first file
bool *m_pbEmptyClms2; // .... the second file
bool *m_pbGreenClms1; // 1D array indicating whether a column has its "lookalike" in the second file
bool *m_pbGreenClms2; // 1D array indicating whether a column has its "lookalike" in the first file
long *m_pnFoundDifferences; // number of differences found between intersected columns (for the doubleclicked cell)
//SimilaritiesAcrossTables *similaritiesAcrossTables; // the best similarity for each column across tables
std::vector<SimilaritiesAcrossTables> m_vecSimilaritiesAcrossTables;
std::vector<SimilaritiesAcrossTables> m_vecSimilaritiesAcrossTablesSorted;
long m_nSelectedDifference; // difference picked by user in the drop down box in the "analysis" tab
CMFCRibbonBar* m_pRibbon; // pointer to ribbon object
						//CMFCRibbonStatusBarPane *statusBarPane;
Table m_Table1; // Important information of tables (number of columns, rows, first row under header, etc.)
Table m_Table2;
COleSafeArray m_saRet1; // OLE object for connection to first Excel file
COleSafeArray m_saRet2; // ... second Excel file
COleSafeArray m_saTmpRet1; // temporary OLE object - this workaround hopefully protects against cache collisions
COleSafeArray m_saTmpRet2;
CString m_szFilename1; // name of the first file that is to be compared
CString m_szFilename2; // .... second ....
CWorkbooks m_Books1; // TypeLib objects
CWorkbook m_Book1; // ...
CWorksheets m_Sheets1; // ...
CWorksheet m_Sheet1; // ...
CRange m_oRange1; // ...
CWorkbooks m_Books2; // TypeLib objects
CWorkbook m_Book2; // ...
CWorksheets m_Sheets2; // ...
CWorksheet m_Sheet2; // ...
CRange m_oRange2; // ...
CCellFormat m_CellFormat; // TypeLib object
Cnterior m_Interior; // TypeLib object
COleVariant
covTrue((short)TRUE),
covFalse((short)FALSE),
covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR); // OLE constants
CApplication m_App; // application object
CMap <CString, LPCTSTR, long, long> m_Map1; // map for keys in the first file
CMap <CString, LPCTSTR, long, long> m_Map2; // ... second file
										  //CMap <CString, LPCTSTR, long, long> tmpMap1; // map for keys in the first file
										  //CMap <CString, LPCTSTR, long, long> tmpMap2; // ... second file
std::map<CString, long> m_mapTmpMap1; // searching for appropriate keys
std::map<CString, long> m_mapTmpMap2;
int m_nUiToBeRefreshed; // how many times the UI is to be refreshed (just a workaround)
float m_fZoom; // not used at the moment
int m_nPrgval1; // not used 
CMFCRibbonProgressBar* m_pProgressBar1; // CMFCRibbon UI objects
CMFCRibbonProgressBar* m_pProgressBar2;
CMFCRibbonProgressBar* m_pKeyProgressBar1;
CMFCRibbonProgressBar* m_pKeyProgressBar2;
CMFCRibbonComboBox* m_pCombo2;  // Pointers to GUI elements
CMFCRibbonComboBox* m_pSheetCombo1;
CMFCRibbonComboBox* m_pSheetCombo2;
CMFCRibbonEdit* m_pSpinner1_Fdata;
CMFCRibbonEdit* m_pSpinner1_Names;
CMFCRibbonEdit* m_pSpinner2_Fdata;
CMFCRibbonEdit* m_pSpinner2_Names;
CMFCRibbonCheckBox* m_pMarkIn1;
CMFCRibbonCheckBox* m_pMarkIn2;
CMFCRibbonSlider* m_pSlider;
CMFCRibbonButton* m_pUnhideExcel;
CMFCRibbonCheckBox* m_pVerifyKeys;
CMFCRibbonCheckBox* m_pSameNames;
CMFCRibbonColorButton* m_pColorPicker1;
CMFCRibbonColorButton* m_pColorPicker2;
CMFCRibbonCheckBox* m_pAuto;
CMFCRibbonComboBox* m_pFoundDifferences;
CMFCRibbonLabel* m_pLabel0;
CMFCRibbonLabel* m_pLabel1;
CMFCRibbonLabel* m_pLabel2;
CMFCRibbonCheckBox* m_pToFront;
CMFCRibbonCheckBox* m_pShowSims;
CMFCRibbonButton* m_pCreateNewKeys;
CMFCRibbonButton* m_pButton2;
CMFCRibbonCheckBox* m_pUseIndices;
CMFCRibbonEdit* m_pRows1;
CMFCRibbonEdit* m_pCols1;
CMFCRibbonEdit* m_pRows2;
CMFCRibbonEdit* m_pCols2;
bool m_bToFront; // should Excel be moved to front when the difference is requested to be shown?
int m_nScrolled_X; // how many cells did we scroll horizontally?
int m_nScrolled_Y; // how many cells did we scroll vertically?
ChosenCell M_CCell; // this structure contains coordinates of the cell the mouse pointer is hovering above.
ChosenCell m_CClickedCell;
ChosenCell m_CPrevClickedCell;
ChosenCell m_OldCell;
VisTopLeft m_VisTopLeft; // the coordinates of the topmost and leftmost visible cell
bool m_bIn1file; // whether are differences to be marked in the first file
bool m_bIn2file; // ... second file
bool m_bToDisplaySimilarClms; // whether similar columns across the tables are to be displayed
bool m_bXSimilarityComputed; // whether we have results of similarity across tables
bool m_bAutoMark; // do we request automatic marking of differences?
bool m_bVerifyKeys; // not used anymore - the check of the uniqueness of keys is mandatory and as such it is accomplished automatically
bool m_bToInitSB; // not used anymore
int m_nCellWidth;   // Cell width in pixels
int m_nCellHeight;  // Cell height in pixels
int m_nRibbonWidth; // Ribbon width in pixels
int m_nViewWidth;   // Workspace width in pixels
int m_nViewHeight;  // Workspace height in pixels
int m_nHScrollPos;  // Horizontal scroll position
int m_nVScrollPos;  // Vertical scroll position
int m_nHPageSize;   // Horizontal page size
int m_nVPageSize;   // Vertical page size
bool m_bOnlyPcnt; // whether we want to see detail of a hovered cell
bool m_bForceNotOnlyPcnt; // inverse of the above (just a helper)
int m_nSldr; // value set on the slider in the "analysis" tab
CPen m_SimsPens[256];
CPen m_KeyCurvePen;
bool m_bUseIndexes;
bool m_bNewFile1, m_bNewFile2;
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
UINT SuggestKeys1ThreadProc(LPVOID pParam);
UINT SuggestKeys2ThreadProc(LPVOID pParam);
UINT MutualCheckThreadProc(LPVOID pParam);
UINT FindSimsThreadProc(LPVOID pParam);
UINT FindSimsThreadProc1(LPVOID pParam);
UINT FindSimsThreadProc2(LPVOID pParam);
CString m_szRsrcs;
// </Declaration of threads>
CChildView::CChildView()
{
	//threadCnt = 1;
	m_szRsrcs = L"";
	m_bToFront = false;
	m_nSelectedDifference = 0;
	m_bForceNotOnlyPcnt = true;
	m_bPrereq1valid = false;
	m_bPrereq2valid = false;
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
	m_pchMainArr1 = new char[1]; // this 2D array contains first character of content of each cell taken from the first XLS file (first horizontally, then vertically)
	m_pchMainArr2 = new char[1]; // .... the second XLS file
	m_pbMarkIn1Arr = new bool[1]; // this 2D array indicates whether a cell at its coordinates (see above) is to be marked in the first file
	m_pbMarkIn2Arr = new bool[1]; // ... in the second file.
	m_pszKeyArr11 = new CString[1]; // array of the strings found in the first key column in the first file
							   //keyArr12 = new CString[1]; // ... the second key ... the first file
							   //keyArr13 = new CString[1]; // ... the third key ... the first file
	m_pszKeyArr21 = new CString[1]; // ... the first key ... the second file
							   //keyArr22 = new CString[1]; // ... the second key ... the second file
							   //keyArr23 = new CString[1]; // ... the third key ... the second file
	m_pszTmpKeyArr11 = new CString[1];
	m_pszTmpKeyArr21 = new CString[1];
	m_pbKeyMissing1 = new bool[1];
	m_pbKeyMissing2 = new bool[1];
	m_pbTmpKeyMissing1 = new bool[1];
	m_pbTmpKeyMissing2 = new bool[1];
	m_pnMainMatrix = new int[1]; // 2D array representing the result matrix
	m_pbMarkedMatrix = new bool[1]; // 2D array indicating marked cells in the result matrix
	m_pbEmptyClms1 = new bool[1]; // 1D array indicating empty columns in the first file
	m_pbEmptyClms2 = new bool[1]; // .... the second file
	m_pbGreenClms1 = new bool[1]; // 1D array indicating whether a column has its "lookalike" in the second file
	m_pbGreenClms2 = new bool[1]; // 1D array indicating whether a column has its "lookalike" in the first file
	m_pnFoundDifferences = new long[1]; // number of differences found between intersected columns (for the doubleclicked cell) 
	//similaritiesAcrossTables = new SimilaritiesAcrossTables[1];
	m_bIn1file = false;
	m_bIn2file = false;
	m_bSameNames = false;
	m_bOnlyPcnt = false;
	m_bToInitSB = true;
	m_VisTopLeft.left = 0;
	m_VisTopLeft.top = 0;
	m_BestKeyComb.pk1 = 0;
	m_BestKeyComb.pk2 = 0;
	m_BestKeyComb.rating = 0;
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
	m_bNewFile1 = false;
	m_bNewFile2 = false;
	m_nComplexity = 100000;
}
CChildView::~CChildView()
{
	delete[] m_pchMainArr1;
	delete[] m_pchMainArr2; //char[1]; // ....  XLS files
	delete[] m_pbMarkIn1Arr; //bool[1]; // this 2D array indicates whether a cell at its coordinates (see above) is to be marked in the first file
	delete[] m_pbMarkIn2Arr; //bool[1]; // ... in the second file.
	delete[] m_pszKeyArr11; //CString[1]; // array of the strings found in the first key column in the first file
							   //keyArr12; //CString[1]; // ... the second key ... the first file
							   //keyArr13; //CString[1]; // ... the third key ... the first file
	delete[] m_pszKeyArr21; //CString[1]; // ... the first key ... the second file
							   //keyArr22; //CString[1]; // ... the second key ... the second file
							   //keyArr23; //CString[1]; // ... the third key ... the second file
	delete[] m_pszTmpKeyArr11; //CString[1];
	delete[] m_pszTmpKeyArr21; //CString[1];
	delete[] m_pbKeyMissing1; //bool[1];
	delete[] m_pbKeyMissing2; //bool[1];
	delete[] m_pbTmpKeyMissing1; //bool[1];
	delete[] m_pbTmpKeyMissing2; //bool[1];
	delete[] m_pnMainMatrix; //int[1]; // 2D array representing the result matrix
	delete[] m_pbMarkedMatrix; //bool[1]; // 2D array indicating marked cells in the result matrix
	delete[] m_pbEmptyClms1; //bool[1]; // 1D array indicating empty columns in the first file
	delete[] m_pbEmptyClms2; //bool[1]; // .... the second file
	delete[] m_pbGreenClms1; //bool[1]; // 1D array indicating whether a column has its "lookalike" in the second file
	delete[] m_pbGreenClms2; //bool[1]; // 1D array indicating whether a column has its "lookalike" in the first file
	delete[] m_pnFoundDifferences; //long[1]; // number of differences found between intersected columns (for the doubleclicked cell) 
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
	int bnd_X_min = 1;
	int bnd_X_max = m_Table2.NumberOfColumns;
	int bnd_Y_min = 1;
	int bnd_Y_max = m_Table1.NumberOfColumns; // ?????
											//visTopLeft.left = 3;
	bool cursor = false;
	// Info area redrawing
	dc.SelectObject(&font4);
	CString prcnt;
	prcnt = L"";
	if (!m_bToDisplaySimilarClms && M_CCell.x * M_CCell.y && m_nNatrixDone)
	{
		dc.SetBkMode(TRANSPARENT);
		if (M_CCell.x <= bnd_X_max && M_CCell.y <= bnd_Y_max)
		{
			long sameness = mxGet(M_CCell.x, M_CCell.y);
			//dc.SelectObject(&pen8);
			if (m_Table1.Columns[M_CCell.y] == m_Table2.Columns[M_CCell.x] && sameness < m_nEffMax)
				dc.SetTextColor(RGB(255, 0, 0));
			else
				dc.SetTextColor(RGB(0, 0, 0));
			prcnt.Format(L"Δ:%i", m_nEffMax - sameness);
			dc.TextOutW(5, 20, prcnt);
			dc.SetTextColor(RGB(0, 255, 0));
			prcnt.Format(L"=:%i", sameness);
			dc.TextOutW(5, 50, prcnt);
			if (m_pbEmptyClms1[M_CCell.y] && m_pbEmptyClms2[M_CCell.x])
			{
				dc.SetTextColor(RGB(0, 0, 0));
				dc.TextOutW(5, 80, CMsg(IDS_EMPTY)); // CMsg(IDS_EMPTY)
			}
		}
	}
	if (!m_bOnlyPcnt)
	{
		if (m_bNewFile1 && m_szFilename1)
		{
			dc.SetTextColor(RGB(120, 0, 130));
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(&font1C);
			int index = ReverseFind(m_szFilename1, L"\\", -1) + 1;
			dc.TextOutW(2, 114, m_szFilename1.Mid(index, 22));
		}
		if (m_bNewFile2 && m_szFilename2)
		{
			dc.SetTextColor(RGB(120, 0, 130));
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(&font2C);
			int index = ReverseFind(m_szFilename2, L"\\", -1) + 1;
			dc.TextOutW(112, 118, m_szFilename2.Mid(index, 22));
		}
	}
	if (M_CCell.x * M_CCell.y && m_bToDisplaySimilarClms)
	{
		if (M_CCell.x <= bnd_X_max && M_CCell.y <= bnd_Y_max)
		{
			dc.SetTextColor(RGB(50, 100, 250));
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(&font1C);
			dc.TextOutW(5, 30, CMsg(IDS_KEY_SUITABILITY));// IDS_KEY_SUITABILITY
			dc.SelectObject(&font4);
			prcnt.Format(L"~ %i%%", 100 * m_vecSimilaritiesAcrossTables[M_CCell.y].similarity / min(m_Table1.NumberOfRows - m_Table1.FirstRowWithData + 1, m_Table2.NumberOfRows - m_Table2.FirstRowWithData + 1));
			dc.TextOutW(15, 60, prcnt);
		}
	}
	// /Info area redrawing
											//dc.TextOutW(100, 100, L"+ěščřžýá"); // test of the ability to make an output in czech lang.
	dc.SelectObject(&brush0);
	int mx_x_adj, mx_y_adj; // for access to data
	dc.SelectObject(&pen2);
	dc.SelectObject(&brush1);
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
	dc.SelectObject(&pen2);
	dc.SelectObject(&brush1);
	//dc.SelectObject(&font1);
	for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max; mx_y++)
	{
		cursor = false;
		mx_y_adj = mx_y + m_VisTopLeft.top;
		dc.SetBkMode(OPAQUE);
		if (isThisAKey(1, mx_y_adj))
		{
			if (mx_y_adj == m_OldCell.y)
			{
				dc.SelectObject(&brush6);
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
						dc.SelectObject(&brush0);
					}
				}
				else
				{
					dc.SelectObject(&brush6);
				}
			}
		}
		else
		{
			dc.SelectObject(&brush0);
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
					dc.SelectObject(&brush6);
				}
				else
				{
					dc.SelectObject(&brush0);
				}
			}
		}
		dc.SelectObject(&pen2);
		//else
			dc.Rectangle(0, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X, 1 + OFFSET_Y + mx_y * STEP_Y + STEP_Y);
		if (cursor)
		{
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(&brush0);
			if (m_bToDisplaySimilarClms) dc.SelectObject(&pen12); else dc.SelectObject(&pen11);
			dc.Rectangle(2, 2 + OFFSET_Y + mx_y * STEP_Y, OFFSET_X + STEP_X - 1, -1 + OFFSET_Y + mx_y * STEP_Y + STEP_Y);
			dc.SetBkMode(OPAQUE);
			dc.SelectObject(&brush0);
			dc.SelectObject(&pen4);
		}
		if (m_nNatrixDone && !m_bOnlyPcnt && ((mx_y - m_VisTopLeft.top) > 0))
		{
			if (m_pbGreenClms1[mx_y])
			{
				dc.SelectObject(&pen5);
				dc.SelectObject(&brush1);
				dc.Ellipse(OFFSET_X, OFFSET_X + (mx_y - m_VisTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1, OFFSET_Y + STEP_Y + (mx_y - m_VisTopLeft.top) * STEP_Y);
			}
			if (m_pbEmptyClms1[mx_y])
			{
				dc.SelectObject(&pen6);
				dc.SelectObject(&brush2);
				dc.Ellipse(OFFSET_X, OFFSET_X + (mx_y - m_VisTopLeft.top) * STEP_Y, OFFSET_X + STEP_X - 1, OFFSET_Y + STEP_Y + (mx_y - m_VisTopLeft.top) * STEP_Y);
			}
		}
		dc.SetBkMode(TRANSPARENT);
		if (isThisAKey(1, mx_y_adj))
		{
			dc.SelectObject(&font1B);
			dc.SetTextColor(RGB(0, 0, 170));
		}
		else
		{
			dc.SelectObject(&font1);
			dc.SetTextColor(RGB(0, 0, 0));
		}
		dc.TextOutW(2, OFFSET_Y + 5 + mx_y * STEP_Y, m_Table1.Columns[mx_y_adj]);
	}
	//dc.SelectObject(&font2);
	for (int mx_x = bnd_X_min; mx_x <= bnd_X_max; mx_x++)
	{
		cursor = false;
		mx_x_adj = mx_x + m_VisTopLeft.left;
		dc.SetBkMode(OPAQUE);
		if (isThisAKey(2, mx_x_adj))
		{
			if (mx_x_adj == m_OldCell.x)
			{
				dc.SelectObject(&brush6);
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
						dc.SelectObject(&brush0);
					}
				}
				else
				{
					dc.SelectObject(&brush6);
				}
			}
		}
		else
		{
			dc.SelectObject(&brush0);
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
					dc.SelectObject(&brush6);
				}
				else
				{
					dc.SelectObject(&brush0);
				}
			}
		}
		dc.SelectObject(&pen2);
		//else
			dc.Rectangle(OFFSET_X + mx_x * STEP_X, 0, 1 + OFFSET_X + mx_x * STEP_X + STEP_X, 1 + OFFSET_Y + STEP_Y);
		if (cursor)
		{
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(&brush0);
			if (m_bToDisplaySimilarClms) dc.SelectObject(&pen12); else dc.SelectObject(&pen11);
			dc.Rectangle(2 + OFFSET_X + mx_x * STEP_X, 2, -1 + OFFSET_X + mx_x * STEP_X + STEP_X, OFFSET_Y + STEP_Y - 1);
			dc.SetBkMode(OPAQUE);
			dc.SelectObject(&brush0);
			dc.SelectObject(&pen4);
		}
		if (m_nNatrixDone && !m_bOnlyPcnt && ((mx_x - m_VisTopLeft.left) > 0))
		{
			if (m_pbGreenClms2[mx_x])
			{
				dc.SelectObject(&pen5);
				dc.SelectObject(&brush1);
				dc.Ellipse(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y, OFFSET_X + STEP_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
			}
			if (m_pbEmptyClms2[mx_x])
			{
				dc.SelectObject(&pen6);
				dc.SelectObject(&brush2);
				dc.Ellipse(OFFSET_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y, OFFSET_X + STEP_X + (mx_x - m_VisTopLeft.left) * STEP_X, OFFSET_Y + STEP_Y - 1);
			}
		}
		dc.SetBkMode(TRANSPARENT);
		if (isThisAKey(2, mx_x_adj))
		{
			dc.SelectObject(&font2B);
			dc.SetTextColor(RGB(0, 0, 170));
		}
		else
		{
			dc.SelectObject(&font2);
			dc.SetTextColor(RGB(0, 0, 0));
		}
		dc.TextOutW(OFFSET_X + 5 + mx_x * STEP_X, -2 + OFFSET_Y + STEP_Y, m_Table2.Columns[mx_x_adj]);
	}
	dc.SelectObject(&pen2);
	if (m_nNatrixDone && !m_bOnlyPcnt)
	{
		dc.SetBkMode(OPAQUE);
		dc.SelectObject(&font3);
		int valSimil;
		CString strSimil;
		if (m_nEffMax)
		{
			for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max - m_VisTopLeft.top; mx_y++)
			{
				for (int mx_x = bnd_X_min; mx_x <= bnd_X_max - m_VisTopLeft.left; mx_x++)
				{
					dc.SelectObject(&pen2);
					mx_y_adj = mx_y + m_VisTopLeft.top;
					mx_x_adj = mx_x + m_VisTopLeft.left;
					valSimil = mxGet(mx_x_adj, mx_y_adj) * 100 / m_nEffMax;
					strSimil.Format(L"%i", valSimil);
					strSimil += L"%";
					dc.SetBkMode(OPAQUE);
					if (!m_bSameNames || (m_Table1.Columns[mx_y_adj] == m_Table2.Columns[mx_x_adj]))
					{
						if (valSimil == 100)
						{
							if (m_pbEmptyClms1[mx_y_adj] || m_pbEmptyClms2[mx_x_adj])
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
							if (valSimil > m_nSldr)
							{
								dc.SelectObject(&brush4);
							}
							else
							{
								if (isThisAKey(1, mx_y_adj) || isThisAKey(2, mx_x_adj))
								{
									dc.SelectObject(&brush6);
								}
								else
								{
									dc.SelectObject(&brush0);
								}
							}
						}
					}
					else
					{
						if (isThisAKey(1, mx_y_adj) || isThisAKey(2, mx_x_adj))
						{
							dc.SelectObject(&brush6);
						}
						else
						{
							dc.SelectObject(&brush0);
						}
					}
					if (mx_y_adj == m_CClickedCell.y && mx_x_adj == m_CClickedCell.x)
					{
						dc.SelectObject(&brush6);
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
					if (m_bToDisplaySimilarClms && m_vecSimilaritiesAcrossTables[mx_y_adj].clm2 == mx_x_adj)
					{
						dc.SetBkMode(TRANSPARENT);
						dc.SelectObject(&m_KeyCurvePen);
						dc.SelectObject(&brush7);
						dc.Rectangle(OFFSET_X + (mx_x)* STEP_X + 1, OFFSET_Y + (mx_y)* STEP_Y + 1, OFFSET_X + STEP_X + (mx_x)* STEP_X, OFFSET_Y + STEP_Y + (mx_y)* STEP_Y);
					}
					dc.SetTextColor(RGB(0, 0, 0));
					dc.TextOutW(OFFSET_X + mx_x * STEP_X + 1, OFFSET_Y + mx_y * STEP_Y + 7, strSimil);
				}
			}
			dc.SetBkMode(TRANSPARENT);
			dc.SelectObject(GetStockObject(NULL_BRUSH));
			dc.SelectObject(&pen3);
			for (int mx_y = bnd_Y_min; mx_y <= bnd_Y_max - m_VisTopLeft.top; mx_y++)
			{
				for (int mx_x = bnd_X_min; mx_x <= bnd_X_max - m_VisTopLeft.left; mx_x++)
				{
					mx_y_adj = mx_y + m_VisTopLeft.top;
					mx_x_adj = mx_x + m_VisTopLeft.left;
					if (m_Table1.Columns[mx_y_adj] == m_Table2.Columns[mx_x_adj])
					{
						dc.Rectangle(OFFSET_X + mx_x * STEP_X, OFFSET_Y + mx_y * STEP_Y, 1 + OFFSET_X + STEP_X + mx_x * STEP_X, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y);
					}
					if (mx_y_adj == m_nOldy && mx_x_adj == m_nOldx)
					{
						dc.SelectObject(&pen9);
						dc.Rectangle(OFFSET_X + mx_x * STEP_X + 3, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y - 4, 1 + OFFSET_X + STEP_X + mx_x * STEP_X - 2, 1 + OFFSET_Y + STEP_Y + mx_y * STEP_Y - 2);
						dc.SelectObject(&pen3);
					}
				}
			}
			dc.SelectObject(&pen2);
		}
		m_bOnlyPcnt = false;
	}
	if (m_bToDisplaySimilarClms)
	{
		int mx_x, mx_y = 0;
		//long maxHit = min(table1.NumberOfRows - table1.FirstRowWithData + 1, table2.NumberOfRows - table2.FirstRowWithData + 1);
		long maxHit = m_vecSimilaritiesAcrossTablesSorted[1].similarity;
		for (int s_i = m_vecSimilaritiesAcrossTablesSorted[0].similarityOrder; s_i >= 0; s_i--)
		{
			//ASSERT(s_i != 1);
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
			//ASSERT(mx_y != 9);
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
	dc.SelectObject(&pen4);
	dc.MoveTo(0, OFFSET_Y + STEP_Y);
	dc.LineTo(m_Clnt.w, OFFSET_Y + STEP_Y);
	dc.MoveTo(OFFSET_X + STEP_X, 0);
	dc.LineTo(OFFSET_X + STEP_X, m_Clnt.h);
	m_bOnlyPcnt = false;
}
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
	if (m_Book1)
	{
		try {
			m_Book1.Close(covOptional, covOptional, covOptional);
		}
		catch (COleException* e)
		{
		}
	}
	if (!(CString(fileName) == L""))
	{
		for (int i = 0; i < 255; i++)
		{
			m_Table1.Columns[i] = "";
		}
		if (m_Sheets1 = GetWorksheets1(fileName))
		{
			m_pSheetCombo1->RemoveAllItems();
			for (int i = 1; i <= m_Sheets1.get_Count(); i++)
			{
				if (CWorksheet tempSheet = m_Sheets1.get_Item(COleVariant((short)i)))
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
		mxClear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
		m_nNatrixDone = 0;
		m_OldCell.x = 0;
		m_OldCell.y = 0;
	}
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED)); // CMsg(IDS_FILE_SUCCESFULLY_LOADED)
	deleteAllKeys();
	this->Invalidate();
}
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
	if (m_Book2)
	{
		try {
			m_Book2.Close(covOptional, covOptional, covOptional);
		}
		catch (COleException* e)
		{
		}
	}
	if (!(CString(fileName) == L""))
	{
		for (int i = 0; i < 255; i++)
		{
			m_Table2.Columns[i] = "";
		}
		if (m_Sheets2 = GetWorksheets2(fileName))
		{
			m_pSheetCombo2->RemoveAllItems();
			for (int i = 1; i <= m_Sheets2.get_Count(); i++)
			{
				if (CWorksheet tempSheet = m_Sheets2.get_Item(COleVariant((short)i)))
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
		mxClear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
		m_nNatrixDone = 0;
	}
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_FILE_SUCCESFULLY_LOADED)); // CMsg(IDS_FILE_SUCCESFULLY_LOADED)
	deleteAllKeys();
	this->Invalidate();
}
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
	HWND hWnd0 = this->GetSafeHwnd();
	AfxBeginThread(CreateKeys1ThreadProc, hWnd0);
	AfxBeginThread(CreateKeys2ThreadProc, hWnd0);
	this->Invalidate();
	m_App.put_Visible(true);
	m_App.put_UserControl(TRUE);
}
void CChildView::OnUpdatePickFirstSheet(CCmdUI *pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSheetCombo1 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_PICK_FIRST_SHEET));
}
void CChildView::OnUpdateCreateMatrix(CCmdUI *pCmdUI)
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
void CChildView::OnUpdateFilename1(CCmdUI *pCmdUI)
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
void CChildView::OnUpdateFilename2(CCmdUI *pCmdUI)
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
void CChildView::OnUpdatePickSecondSheet(CCmdUI *pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSheetCombo2 = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_PICK_SECOND_SHEET));
}
void CChildView::OnUpdateProgress1(CCmdUI *pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_PROGRESS2));
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
CWorksheets CChildView::GetWorksheets1(CString TempBookName)
{
	if (!m_App)
	{
		if (!m_App.CreateDispatch(TEXT("Excel.Application")))
		{
			AfxMessageBox(CMsg(IDS_EXCEL_CANNOT_RUN)); // CMsg(IDS_EXCEL_CANNOT_RUN)
			return NULL;
		}
	}
	m_Books1 = m_App.get_Workbooks();
	m_Book1 = m_Books1.Open(TempBookName, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);
	m_App.put_Visible(TRUE);
	m_App.put_UserControl(TRUE);
	return m_Book1.get_Worksheets();
}
CWorksheets CChildView::GetWorksheets2(CString TempBookName)
{
	if (!m_App)
	{
		if (!m_App.CreateDispatch(TEXT("Excel.Application")))
		{
			AfxMessageBox(CMsg(IDS_EXCEL_CANNOT_RUN)); // CMsg(IDS_EXCEL_CANNOT_RUN)
			return NULL;
		}
	}
	m_Books2 = m_App.get_Workbooks();
	m_Book2 = m_Books2.Open(TempBookName, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);
	m_App.put_Visible(TRUE);
	m_App.put_UserControl(TRUE);
	return m_Book2.get_Worksheets();
}
void CChildView::OnPickFirstSheet()
{
	int tmpWSN = m_pSheetCombo1->GetCurSel() + 1;
	CString tmpWSS = m_pSheetCombo1->GetEditText();
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_PRELIM_CHK)); // CMsg(IDS_WAIT_PRELIM_CHK)
	if (tmpWSN > 0)
	{
		m_saRet1.Destroy();
		m_Table1.WorkSheetNumber = tmpWSN;
		m_Sheet1 = m_Sheets1.get_Item(COleVariant(tmpWSS));
		m_oRange1 = m_Sheet1.get_UsedRange();
		m_saRet1 = m_oRange1.get_Value(covOptional);
		long iRows;
		long iCols;
		m_saRet1.GetUBound(1, &iRows);
		m_saRet1.GetUBound(2, &iCols);
		m_Table1.MaxNumberOfRows = iRows;
		m_Table1.MaxNumberOfCols = iCols;
		m_Table1.NumberOfColumns = iCols;
		m_Table1.NumberOfRows = iRows;
		m_Table1.RowWithNames = 1;
		m_saTmpRet1.Destroy();
		m_saTmpRet1 = m_oRange1.get_Value(covOptional);
		m_saTmpRet1.GetUBound(1, &iRows);
		m_saTmpRet1.GetUBound(2, &iCols);
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
			mxClear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
			m_nNatrixDone = 0;
			m_OldCell.x = 0;
			m_OldCell.y = 0;
			M_CCell.x = 0;
			M_CCell.y = 0;
		}
		HWND hWnd0 = this->GetSafeHwnd();
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED)); // CMsg(IDS_DATA_VERIFIED)
		AfxBeginThread(makePrereq1ThreadProc, hWnd0);
	}
}
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
void CChildView::OnUpdateSpin1Names(CCmdUI *pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner1_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN1_NAMES));
}
void CChildView::OnUpdateSpin1Fdata(CCmdUI *pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner1_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN1_FDATA));
}
void CChildView::OnSpin1Fdata()
{
	CString tmps = m_pSpinner1_Fdata->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 2) tmpi = 1; 
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	m_pSpinner1_Fdata->SetEditText(tmps);
	m_Table1.FirstRowWithData = tmpi;
	m_bPrereq1valid = false;
}
void CChildView::updateCombos1()
{
	long index[2];
	CString szdata;
	COleVariant vData;
	for (int i = 1; i <= m_Table1.NumberOfColumns; i++)
	{
		// Loop through the data and report the contents.
		index[0] = m_Table1.RowWithNames;
		index[1] = i;
		try {
			m_saRet1.GetElement(index, vData); vData = (CString)vData;
		}
		catch (COleException* e)
		{
			vData = L"";
		}
		szdata = vData;
		if (szdata == "") szdata = CMsg(IDS_NO_NAME); // CMsg(IDS_NO_NAME)
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
void CChildView::OnPickSecondSheet()
{
	int tmpWSN = m_pSheetCombo2->GetCurSel() + 1;
	CString tmpWSS = m_pSheetCombo2->GetEditText();
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)); // CMsg(IDS_WAIT_UNTIL_PRELIMINARY_CHECK)
	if (tmpWSN > 0)
	{
		m_saRet2.Destroy();
		m_Table2.WorkSheetNumber = tmpWSN;
		m_Sheet2 = m_Sheets2.get_Item(COleVariant(tmpWSS));
		m_oRange2 = m_Sheet2.get_UsedRange();
		m_saRet2 = m_oRange2.get_Value(covOptional);
		long iRows;
		long iCols;
		m_saRet2.GetUBound(1, &iRows);
		m_saRet2.GetUBound(2, &iCols);
		m_Table2.MaxNumberOfRows = iRows;
		m_Table2.MaxNumberOfCols = iCols;
		m_Table2.NumberOfColumns = iCols;
		m_Table2.NumberOfRows = iRows;
		m_Table2.RowWithNames = 1;
		m_saTmpRet2.Destroy();
		m_saTmpRet2 = m_oRange2.get_Value(covOptional);
		m_saTmpRet2.GetUBound(1, &iRows);
		m_saTmpRet2.GetUBound(2, &iCols);
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
			mxClear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
			m_nNatrixDone = 0;
			m_OldCell.x = 0;
			m_OldCell.y = 0;
			M_CCell.x = 0;
			M_CCell.y = 0;
		}
		updateCombos2();
		this->Invalidate();
		m_nNatrixDone = false;
		HWND hWnd0 = this->GetSafeHwnd();
		if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_DATA_VERIFIED)); // CMsg(IDS_DATA_VERIFIED)
		AfxBeginThread(makePrereq2ThreadProc, hWnd0);
	}
}
void CChildView::updateCombos2()
{
	long index[2];
	CString szdata;
	COleVariant vData;
	for (int i = 1; i <= m_Table2.NumberOfColumns; i++)
	{
		index[0] = m_Table2.RowWithNames;
		index[1] = i;
		try {
			m_saRet2.GetElement(index, vData); vData = (CString)vData;
		}
		catch (COleException* e)
		{
			vData = L"";
		}
		szdata = vData;
		if (szdata == "") szdata = CMsg(IDS_NO_NAME); // CMsg(IDS_NO_NAME)
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
void CChildView::OnUpdateSpin2Fdata(CCmdUI *pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner2_Fdata = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN2_FDATA));
}
void CChildView::OnSpin2Fdata()
{
	CString tmps = m_pSpinner2_Fdata->GetEditText();
	int tmpi = _ttoi(tmps);
	if (tmpi < 2) tmpi = 1;
	if (tmpi > 64) tmpi = 64;
	tmps.Format(_T("%d"), tmpi);
	m_pSpinner2_Fdata->SetEditText(tmps);
	m_Table2.FirstRowWithData = tmpi;
	m_bPrereq2valid = false;
}
void CChildView::OnUpdateSpin2Names(CCmdUI *pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSpinner2_Names = DYNAMIC_DOWNCAST(CMFCRibbonEdit, m_pRibbon->FindByID(ID_SPIN2_NAMES));
}
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
void CChildView::makeCharArr1()
{
	if (int arSize1 = (m_Table1.NumberOfColumns + 1) * (m_Table1.NumberOfRows + 1))
	{
		long prgHlpr0, prgHlpr;
		prgHlpr0 = 0;
		prgHlpr = 0;
		delete[] m_pchMainArr1;
		m_pchMainArr1 = new char[arSize1];
		long index[2];
		char chr;
		COleVariant vData;
		CString szdata;
		for (int i_c = 1; i_c <= m_Table1.NumberOfColumns; i_c++)
		{
			prgHlpr0 = 100 * i_c / m_Table1.NumberOfColumns;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
			}
			for (int i_r = 1; i_r <= m_Table1.NumberOfRows; i_r++)
			{
				index[0] = i_r;
				index[1] = i_c;
				try {
					m_saRet1.GetElement(index, vData); vData = (CString)vData;
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
				m_pchMainArr1[(i_r - 1) * m_Table1.NumberOfColumns + i_c] = chr;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 100);
}
void CChildView::makeCharArr2()
{
	if (int arSize2 = (m_Table2.NumberOfColumns + 1) * (m_Table2.NumberOfRows + 1))
	{
		long prgHlpr0, prgHlpr;
		prgHlpr0 = 0;
		prgHlpr = 0;
		delete[] m_pchMainArr2;
		m_pchMainArr2 = new char[arSize2];
		long index[2];
		char chr;
		COleVariant vData;
		CString szdata;
		for (int i_c = 1; i_c <= m_Table2.NumberOfColumns; i_c++)
		{
			prgHlpr0 = 100 * i_c / m_Table2.NumberOfColumns;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
			}
			//TRACE("i_c: %i\n", i_c);
			for (int i_r = 1; i_r <= m_Table2.NumberOfRows; i_r++)
			{
				//TRACE("i_r: %i, i_c: %i\n", i_r, i_c);
				index[0] = i_r;
				index[1] = i_c;
				try {
					m_saRet2.GetElement(index, vData); vData = (CString)vData;
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
				m_pchMainArr2[(i_r - 1) * m_Table2.NumberOfColumns + i_c] = chr;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
}
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
		HWND hWnd0 = this->GetSafeHwnd();
		int mx_X_max = m_Table2.NumberOfColumns;
		m_pbMarkedMatrix[(M_CCell.y - 1) * mx_X_max + M_CCell.x] = true;
		this->Invalidate();
		AfxBeginThread(MyThreadProc3, hWnd0);
	}
	CWnd::OnLButtonDblClk(nFlags, point);
}
void CChildView::mxClear(int x, int y)
{
	int size = (x + 1) * (y + 1);
	delete[] m_pnMainMatrix;
	m_pnMainMatrix = new int[size];
	delete[] m_pbMarkedMatrix;
	m_pbMarkedMatrix = new bool[size];
	for (int i = 0; i < size; i++)
	{
		m_pnMainMatrix[i] = 0;
		m_pbMarkedMatrix[i] = false;
	}
}
int CChildView::mxPut(int x, int y)
{
	int mx_X_max = m_Table2.NumberOfColumns;
	int index = (y - 1) * mx_X_max + x;
	m_pnMainMatrix[index] += 1;
	return 0;
}
int CChildView::mxGet(int x, int y)
{
	int mx_X_max = m_Table2.NumberOfColumns;
	int index = (y - 1) * mx_X_max + x;
	return m_pnMainMatrix[index];
}
bool CChildView::mxMarkedGet(int x, int y)
{
	int mx_X_max = m_Table2.NumberOfColumns;
	int index = (y - 1) * mx_X_max + x;
	return m_pbMarkedMatrix[index];
}
void CChildView::checkEmptiness1()
{
	delete[] m_pbEmptyClms1;
	m_pbEmptyClms1 = new bool[m_Table1.NumberOfColumns + 2];
	for (int i = 0; i <= m_Table1.NumberOfColumns; i++) m_pbEmptyClms1[i] = true;
	long prgHlpr0, prgHlpr;
	prgHlpr0 = 0;
	prgHlpr = 0;
	for (int i_c = 1; i_c <= m_Table1.NumberOfColumns; i_c++)
	{
		prgHlpr0 = 100 * i_c / m_Table1.NumberOfColumns;
		if (prgHlpr0 > prgHlpr + 10)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
		}
		for (int i_r = m_Table1.FirstRowWithData; i_r <= m_Table1.NumberOfRows; i_r++)
		{
			if (m_pchMainArr1[(i_r - 1) * m_Table1.NumberOfColumns + i_c])
			{
				m_pbEmptyClms1[i_c] = false;
				break;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 100);
}
void CChildView::checkEmptiness2()
{
	delete[] m_pbEmptyClms2;
	m_pbEmptyClms2 = new bool[m_Table2.NumberOfColumns + 2];
	for (int i = 0; i <= m_Table2.NumberOfColumns; i++) m_pbEmptyClms2[i] = true;
	long prgHlpr0, prgHlpr;
	prgHlpr0 = 0;
	prgHlpr = 0;
	for (int i_c = 1; i_c <= m_Table2.NumberOfColumns; i_c++)
	{
		prgHlpr0 = 100 * i_c / m_Table2.NumberOfColumns;
		if (prgHlpr0 > prgHlpr + 10)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
		}
		for (int i_r = m_Table2.FirstRowWithData; i_r <= m_Table2.NumberOfRows; i_r++)
		{
			if (m_pchMainArr2[(i_r - 1) * m_Table2.NumberOfColumns + i_c])
			{
				m_pbEmptyClms2[i_c] = false;
				break;
			}
		}
	}
	PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
}
bool CChildView::checkKeysUniqueness1()
{
	m_bLockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	CString szTaken_A, szTaken_B;
	for (int i0 = m_Table1.FirstRowWithData; i0 <= m_Table1.NumberOfRows; i0++)
	{
		prgHlpr0 = 100 * i0 / m_Table1.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
		}
		szTaken_A = m_pszKeyArr11[i0];
		for (int i1 = i0 + 1; i1 <= m_Table1.NumberOfRows; i1++)
		{
			szTaken_B = m_pszKeyArr11[i1];
			if (szTaken_A == szTaken_B)
			{
				m_bLockPrg1 = false;
				PostMessage(CM_UPDATE_PROGRESS, 0, 100);
				return false;
			}
		}
	}
	m_bLockPrg1 = false;
	return true;
}
bool CChildView::checkKeysUniqueness2()
{
	m_bLockPrg2 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	CString szTaken_A, szTaken_B;
	for (int i0 = m_Table2.FirstRowWithData; i0 <= m_Table2.NumberOfRows; i0++)
	{
		prgHlpr0 = 100 * i0 / m_Table2.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
		}
		szTaken_A = m_pszKeyArr21[i0];
		for (int i1 = i0 + 1; i1 <= m_Table2.NumberOfRows; i1++)
		{
			szTaken_B = m_pszKeyArr21[i1];
			if (szTaken_A == szTaken_B)
			{
				m_bLockPrg2 = false;
				PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
				return false;
			}
		}
	}
	m_bLockPrg2 = false;
	return true;
}
void CChildView::firstPass()
{
	if (!m_bPrereq1valid) makePrereq1();
	if (!m_bPrereq2valid) makePrereq2();
	m_bDoAutoMark = m_bAutoMark;
	m_bLockPrg1 = true;
	CString concatenatedKey1, concatenatedKey2;
	int prgHlpr = 0, prgHlpr0 = 0;
	char firstChar1, firstChar2;
	m_nEffMax = 0;
	mxClear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
	POSITION mapPos1;
	mapPos1 = m_Map1.GetStartPosition();
	//// The commented code below is used only if the keys are stored in arrays instead of maps
	long /*keyRow1, */ keyRow2;
	int fchar1_y, fchar2_y;
	//int i1; // iterator for progress visualisation;
	//i1 = table1.FirstRowWithData-1;
	if (m_bAutoMark)
	{
		for (long i1 = m_Table1.FirstRowWithData; i1 <= m_Table1.NumberOfRows; i1++)
			//while (mapPos1 !=  NULL)
		{
			//i1++;
			prgHlpr0 = 99 * i1 / m_Table1.NumberOfRows; // 99: because 100 would terminate the thread immaturely
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				//pProgressBar1->SetPos(prgHlpr);
			}
			//map1.GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);
			concatenatedKey1 = m_pszKeyArr11[i1];
			//for (int i2 = table2.FirstRowWithData; i2 <= table2.NumberOfRows; i2++)
			//{
			//concatenatedKey2 = keyArr21[i2];
			//if (concatenatedKey1 == concatenatedKey2)
			if (m_Map2.Lookup(concatenatedKey1, (long&)keyRow2))
			{
				m_nEffMax++;
				//procBoundaries[thrdIdx].effPortion++;
				fchar1_y = (i1 - 1) * m_Table1.NumberOfColumns;
				for (int i3 = 1; i3 <= m_Table1.NumberOfColumns; i3++)
				{
					firstChar1 = m_pchMainArr1[fchar1_y + i3];
					fchar2_y = (keyRow2 - 1) * m_Table2.NumberOfColumns;
					for (int i4 = 1; i4 <= m_Table2.NumberOfColumns; i4++)
					{
						firstChar2 = m_pchMainArr2[fchar2_y + i4];
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
								if (m_Table1.Columns[i3] == m_Table2.Columns[i4])
								{
									m_pchMainArr1[fchar1_y + i3] = 1;
									m_pchMainArr2[fchar2_y + i4] = 1;
								}
							}
						}
						else
						{
							if (m_Table1.Columns[i3] == m_Table2.Columns[i4])
							{
								m_pchMainArr1[fchar1_y + i3] = 1;
								m_pchMainArr2[fchar2_y + i4] = 1;
							}
						}
					}
				}
			}
			else
			{
				m_pbKeyMissing1[i1] = true;
			}
			//}
		}
		if (m_bIn2file)
		{
			long keyRow1;
			prgHlpr = 0; prgHlpr0 = 0;
			for (long i1_2 = m_Table2.FirstRowWithData; i1_2 <= m_Table2.NumberOfRows; i1_2++)
			{
				prgHlpr0 = 100 * i1_2 / m_Table2.NumberOfRows;
				if (prgHlpr0 > prgHlpr)
				{
					prgHlpr = prgHlpr0;
					PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				}
				concatenatedKey2 = m_pszKeyArr21[i1_2];
				if (!m_Map1.Lookup(concatenatedKey2, (long&)keyRow1))
				{
					m_pbKeyMissing2[i1_2] = true;
				}
			}
		}
		PostMessage(CM_UPDATE_PROGRESS, 0, 100); // because otherwise the "resolve auto mark" procedure would be started prematurely
	}
	else
	{
		for (long i1 = m_Table1.FirstRowWithData; i1 <= m_Table1.NumberOfRows; i1++)
			//while (mapPos1 !=  NULL)
		{
			//i1++;
			prgHlpr0 = 100 * i1 / m_Table1.NumberOfRows;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				//pProgressBar1->SetPos(prgHlpr);
			}
			//map1.GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);
			concatenatedKey1 = m_pszKeyArr11[i1];
			//for (int i2 = table2.FirstRowWithData; i2 <= table2.NumberOfRows; i2++)
			//{
			//concatenatedKey2 = keyArr21[i2];
			//if (concatenatedKey1 == concatenatedKey2)
			if (m_Map2.Lookup(concatenatedKey1, (long&)keyRow2))
			{
				m_nEffMax++;
				//procBoundaries[thrdIdx].effPortion++;
				fchar1_y = (i1 - 1) * m_Table1.NumberOfColumns;
				for (int i3 = 1; i3 <= m_Table1.NumberOfColumns; i3++)
				{
					firstChar1 = m_pchMainArr1[fchar1_y + i3];
					fchar2_y = (keyRow2 - 1) * m_Table2.NumberOfColumns;
					for (int i4 = 1; i4 <= m_Table2.NumberOfColumns; i4++)
					{
						firstChar2 = m_pchMainArr2[fchar2_y + i4];
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
	delete[] m_pbGreenClms1;
	m_pbGreenClms1 = new bool[m_Table1.NumberOfColumns + 2];
	delete[] m_pbGreenClms2;
	m_pbGreenClms2 = new bool[m_Table2.NumberOfColumns + 2];
	for (int i = 0; i <= m_Table1.NumberOfColumns; i++) m_pbGreenClms1[i] = false; // TODO: could have been moved to the cyclus below in the "else" branch - for better performance
	for (int i = 0; i <= m_Table2.NumberOfColumns; i++) m_pbGreenClms2[i] = false; // dtto
	for (int i_c = 1; i_c <= m_Table2.NumberOfColumns; i_c++)
	{
		for (int i_r = 1; i_r <= m_Table1.NumberOfColumns; i_r++)
		{
			if (mxGet(i_c, i_r) == m_nEffMax)
			{
				m_pbGreenClms1[i_r] = true;
				m_pbGreenClms2[i_c] = true;
			}
		}
	}
	m_nNatrixDone++;
	PostMessage(CM_UPDATE_PROGRESS, 0, 1000);
	m_bLockPrg1 = false;
}
int CChildView::createKeyArrays1()
{
	m_NotUniqueKeys1 = { 0, 0, L"" };
	long mapIdx;
	long index[2];
	COleVariant vData;
	CString szdata;
	long idx = 0;
	m_Map1.RemoveAll();
	CString testdata;
	delete[] m_pszKeyArr11;
	m_pszKeyArr11 = new CString[m_Table1.NumberOfRows + 2];
	delete[] m_pbKeyMissing1;
	m_pbKeyMissing1 = new bool[m_Table1.NumberOfRows + 2];
	m_bLockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	for (int i_i = m_Table1.FirstRowWithData; i_i <= m_Table1.NumberOfRows; i_i++)
	{
		prgHlpr0 = 100 * i_i / m_Table1.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
		}
		szdata = "";
		int nthKey;
		for (int k_i = 0; k_i < m_nKeyPairCounter; k_i++)
		{
			nthKey = getNthKey(1, k_i);
			if (nthKey)
			{
				// Loop through the data and report the contents.
				index[0] = i_i;
				index[1] = nthKey;
				try {
					m_saRet1.GetElement(index, vData); vData = (CString)vData;
				}
				catch (COleException* e)
				{
					vData = L"";
				}
				szdata += vData;
			}
		}
		m_pbKeyMissing1[i_i] = false;
		if (m_bUseIndexes)
		{
			idx = 0;
			do {
				idx++;
				testdata.Format(L"%s_idx%i", szdata, idx);
			} while (m_Map1.Lookup(testdata, (long&)mapIdx));
			szdata = testdata;
		}
		else
		{
			if (m_Map1.Lookup(szdata, (long&)mapIdx))
			{
				m_NotUniqueKeys1 = { i_i, mapIdx, szdata };
				m_Map1.RemoveAll();
				return 1;
			}
		}
		m_pszKeyArr11[i_i] = szdata;
		m_Map1.SetAt(szdata, i_i);
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 1000);
	return 0;
}
int CChildView::createKeyArrays2()
{
	m_NotUniqueKeys2 = { 0, 0, L"" };
	long mapIdx;
	long index[2];
	COleVariant vData;
	CString szdata;
	long idx = 0;
	m_Map2.RemoveAll();
	CString testdata;
	delete[] m_pszKeyArr21;
	m_pszKeyArr21 = new CString[m_Table2.NumberOfRows + 2];
	delete[] m_pbKeyMissing2;
	m_pbKeyMissing2 = new bool[m_Table2.NumberOfRows + 2];
	m_bLockPrg2 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	prgHlpr = 0;
	prgHlpr0 = 0;
	for (int i_i = m_Table2.FirstRowWithData; i_i <= m_Table2.NumberOfRows; i_i++)
	{
		prgHlpr0 = 100 * i_i / m_Table2.NumberOfRows;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
		}
		szdata = "";
		int nthKey;
		for (int k_i = 0; k_i < m_nKeyPairCounter; k_i++)
		{
			nthKey = getNthKey(2, k_i);
			if (nthKey)
			{
				// Loop through the data and report the contents.
				index[0] = i_i;
				index[1] = nthKey;
				try {
					m_saRet2.GetElement(index, vData); vData = (CString)vData;
				}
				catch (COleException* e)
				{
					vData = L"";
				}
				szdata += vData;
			}
		}
		m_pbKeyMissing2[i_i] = false;
		if (m_bUseIndexes)
		{
			idx = 0;
			do {
				idx++;
				testdata.Format(L"%s_idx%i", szdata, idx);
			} while (m_Map2.Lookup(testdata, (long&)mapIdx));
			szdata = testdata;
		}
		else
		{
			if (m_Map2.Lookup(szdata, (long&)mapIdx))
			{
				m_NotUniqueKeys2 = { i_i, mapIdx, szdata };  // this error is on purpose
				m_Map2.RemoveAll();
				return 2;
			}
		}
		m_pszKeyArr21[i_i] = szdata;
		m_Map2.SetAt(szdata, i_i);
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
		m_saRet1.GetElement(index, vData); vData = (CString)vData;
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
		m_saRet2.GetElement(index, vData); vData = (CString)vData;
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
void CChildView::OnUpdateSlider2(CCmdUI *pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSlider = DYNAMIC_DOWNCAST(CMFCRibbonSlider, m_pRibbon->FindByID(ID_SLIDER2));
	if (m_pSlider->GetPos() == 0)
		m_pSlider->SetPos(m_nSldr);
}
void CChildView::OnCheck4()
{
	m_bIn1file = !m_bIn1file;
}
void CChildView::OnUpdateCheck4(CCmdUI *pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pCmdUI->SetCheck(m_bIn1file);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pMarkIn1 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK4));
}
void CChildView::OnCheck5()
{
	m_bIn2file = !m_bIn2file;
}
void CChildView::OnUpdateCheck5(CCmdUI *pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
	pCmdUI->SetCheck(m_bIn2file);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pMarkIn2 = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK5));
}
void CChildView::OnButton2()
{
	if (m_bLockPrg1 || m_bLockPrg2) {
		MessageBox(CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}
/*	if (bestKeyComb.rating)
	{
		MessageBox(L"Vhodná kombinace klíčů již byla nalezena"); // CMsg(IDS_ANOTHER_PROCESS_STILL_RUNNING)
		return;
	}  */
	if (m_Table1.NumberOfColumns * m_Table2.NumberOfColumns)
	{
		m_bWaitingForKeys = true;
		m_bKeysGathering1done = false;
		m_bKeysGathering2done = false;
		clearPossibleKeys();
		HWND hWnd0 = this->GetSafeHwnd();
		AfxBeginThread(SuggestKeys1ThreadProc, hWnd0);
		AfxBeginThread(SuggestKeys2ThreadProc, hWnd0);
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
	CRange range = m_Sheet1.get_Range(COleVariant(cnv), COleVariant(cnv));
	m_Interior = range.get_Interior();
	m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue))));
	return;
}
void CChildView::markIn2(int row, int clm)
{
	CString cnv = convertR1C1(row, clm);
	CRange range = m_Sheet2.get_Range(COleVariant(cnv), COleVariant(cnv));
	m_Interior = range.get_Interior();
	m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue))));
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
	return 0; 
}
void CChildView::OnUpdateProgress2(CCmdUI *pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pProgressBar2 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_PROGRESS3));
	// Emergency update of the container for found differences
	m_pFoundDifferences = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_pRibbon->FindByID(ID_DIFFS_LIST));
	m_pToFront = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_PUT_TO_FRONT));
}
void CChildView::OnUpdateCheck2(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(m_bVerifyKeys);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pVerifyKeys = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK2));
}
void CChildView::OnCheck2()
{
	m_bVerifyKeys = !m_bVerifyKeys;
}
void CChildView::OnUpdateButton2(CCmdUI *pCmdUI)
{
	if (!&m_App) pCmdUI->Enable(false); else pCmdUI->Enable(true);
	//pCmdUI->SetText(m_bUseIndexes ? L"Sestavit klíč" : L"Najít klíč");
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pButton2 = DYNAMIC_DOWNCAST(CMFCRibbonButton, m_pRibbon->FindByID(ID_BUTTON2));
}
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
void CChildView::OnUpdateCheck7(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(m_bSameNames);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pSameNames = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_CHECK7));
}
UINT MyThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
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
					HWND hWnd0 = this->GetSafeHwnd();
					m_bWaitingForKeys = false;
					m_bKeys1done = false;
					m_bKeys2done = false;
					AfxBeginThread(MyThreadProc, hWnd0);
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
					HWND hWnd0 = this->GetSafeHwnd();
					m_bWaitingForKeys = false;
					m_bKeysGathering1done = false;
					m_bKeysGathering2done = false;
					AfxBeginThread(MutualCheckThreadProc, hWnd0);
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
			tmpS.Format(CMsg(IDS_KEY_COMB_FOUND), m_BestKeyComb.cnt); // CMsg(IDS_KEY_COMB_FOUND)
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
					HWND hWnd0 = this->GetSafeHwnd();
					m_bWaitingForKeys = false;
					m_bKeys1done = false;
					m_bKeys2done = false;
					AfxBeginThread(MyThreadProc, hWnd0);
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
					HWND hWnd0 = this->GetSafeHwnd();
					m_bWaitingForKeys = false;
					m_bKeysGathering1done = false;
					m_bKeysGathering2done = false;
					AfxBeginThread(MutualCheckThreadProc, hWnd0);
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
						fndDfrnc1 += getCellValue1(m_nOldy, i1);
						fndDfrnc1 = fndDfrnc1.Left(26);
						fndDfrnc2 = L"";
						fndDfrnc2.Format(L"   (2r%i):", dfrncRow2);
						fndDfrnc2 += getCellValue2(m_nOldx, dfrncRow2);
						fndDfrnc2 = fndDfrnc2.Left(26);
						selKey = L"";
						selKey.Format(L"%s%s   (key): %s", fndDfrnc1, fndDfrnc2, m_pszKeyArr11[i1]);
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
							CRange range = m_Sheet1.get_Range(COleVariant(starts), COleVariant(ends));
							m_Interior = range.get_Interior();
							m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue))));
							starts = L"";
							ends = L"";
						}
					}
				}
			}
			if (m_bIn1file && !(starts == L"") && !(ends == L""))
			{
				CRange range = m_Sheet1.get_Range(COleVariant(starts), COleVariant(ends));
				m_Interior = range.get_Interior();
				m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue))));
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
						CRange range = m_Sheet2.get_Range(COleVariant(starts), COleVariant(ends));
						m_Interior = range.get_Interior();
						m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue))));
						starts = L"";
						ends = L"";
					}
				}
			}
			if (!(starts == L"") && !(ends == L""))
			{
				CRange range = m_Sheet2.get_Range(COleVariant(starts), COleVariant(ends));
				m_Interior = range.get_Interior();
				m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue))));
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
UINT MyThreadProc2(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	m_bUniqueKeys1 = false;
	m_bUniqueKeys2 = false;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;
	rslt = pWnd->createKeyArrays1();
	if (rslt == 1)
	{
		pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE)); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE)
		return 0;
	}
	m_bUniqueKeys1 = true;
	rslt = pWnd->createKeyArrays2();
	if (rslt == 2)
	{
		pWnd->MessageBox(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE));// CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE)
		return 0;
	}
	AfxEndThread(0);
	return 0;
}
UINT CreateKeys1ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	m_bUniqueKeys1 = false;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;
	rslt = pWnd->createKeyArrays1();
	if (rslt == 1)
	{
		CString s;
		NotUniqueKeys* nu;
		nu = &m_NotUniqueKeys1;
		s.Format(CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow, nu->secondRow); // CMsg(IDS_CHOSEN_KEYS1_NOT_UNIQUE_KEYS)
		pWnd->MessageBox(s);
		m_bLockPrg1 = false;
		return 0;
	}
	m_bUniqueKeys1 = true;
	AfxEndThread(0);
	return 0;
}
UINT CreateKeys2ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	m_bUniqueKeys2 = false;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	int rslt;
	rslt = pWnd->createKeyArrays2();
	if (rslt == 2)
	{
		CString s;
		NotUniqueKeys* nu;
		nu = &m_NotUniqueKeys2;
		s.Format(CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS), nu->keyString, nu->firstRow, nu->secondRow); // CMsg(IDS_CHOSEN_KEYS2_NOT_UNIQUE_KEYS)
		pWnd->MessageBox(s);
		m_bLockPrg2 = false;
		return 0;
	}
	m_bUniqueKeys2 = true;
	AfxEndThread(0);
	return 0;
}
UINT makePrereq1ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->makePrereq1();
	AfxEndThread(0);
	return 0;
}
UINT makePrereq2ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->makePrereq2();
	AfxEndThread(0);
	return 0;
}
UINT MyThreadProc3(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->markInFiles();
	AfxEndThread(0);
	return 0;
}
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
	delete[] m_pbMarkIn1Arr;
	m_pbMarkIn1Arr = new bool[m_Table1.NumberOfRows + 2];
	delete[] m_pbMarkIn2Arr;
	m_pbMarkIn2Arr = new bool[m_Table2.NumberOfRows + 2];
	delete[] m_pnFoundDifferences;
	m_pnFoundDifferences = new long[m_Table1.NumberOfRows + 2];
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
	mapPos1 = m_Map1.GetStartPosition();
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
		m_Map1.GetNextAssoc(mapPos1, concatenatedKey1, (long&)keyRow1);
		if (m_Map2.Lookup(concatenatedKey1, (long&)keyRow2))
		{
			if (!(getCellValue1(cy, keyRow1) == getCellValue2(cx, keyRow2)))
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
void CChildView::OnButton5()
{
	COLORREF i = (int)m_pColorPicker1->GetSelectedItem();
	m_nChosenColor1 = i;
}
void CChildView::OnUpdateButton5(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(true);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pColorPicker1 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_pRibbon->FindByID(ID_BUTTON5));
}
void CChildView::OnButton3()
{
	COLORREF i = (int)m_pColorPicker2->GetSelectedItem();
	m_nChosenColor2 = i;
}
void CChildView::OnUpdateButton3(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(true);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pColorPicker2 = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_pRibbon->FindByID(ID_BUTTON3));
}
void CChildView::OnCheck3()
{
	m_bAutoMark = !m_bAutoMark;
}
void CChildView::OnUpdateCheck3(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(true);
	pCmdUI->SetCheck(m_bAutoMark);
}
void CChildView::makePrereq1()
{
	m_bPrereq1valid = false;
	delete[] m_pchMainArr1;
	m_pchMainArr1 = new char[m_Table1.NumberOfRows + 2];
	makeCharArr1();
	checkEmptiness1();
	m_bPrereq1valid = true;
}
void CChildView::makePrereq2()
{
	m_bPrereq2valid = false;
	delete[] m_pchMainArr2;
	m_pchMainArr2 = new char[m_Table2.NumberOfRows + 2];
	makeCharArr2();
	checkEmptiness2();
	m_bPrereq2valid = true;
}
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
						if (m_pchMainArr1[(r1 - 1) * m_Table1.NumberOfColumns + c1] == 1)
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
								CRange range = m_Sheet1.get_Range(COleVariant(starts), COleVariant(ends));
								m_Interior = range.get_Interior();
								m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue))));
								starts = L"";
								ends = L"";
							}
						}
					}
					if (!(starts == L"") && !(ends == L""))
					{
						CRange range = m_Sheet1.get_Range(COleVariant(starts), COleVariant(ends));
						m_Interior = range.get_Interior();
						m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue))));
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
						if (m_pchMainArr2[(r2 - 1) * m_Table2.NumberOfColumns + c2] == 1)
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
								CRange range = m_Sheet2.get_Range(COleVariant(starts), COleVariant(ends));
								m_Interior = range.get_Interior();
								m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue))));
								starts = L"";
								ends = L"";
							}
						}
					}
					if (!(starts == L"") && !(ends == L""))
					{
						CRange range = m_Sheet2.get_Range(COleVariant(starts), COleVariant(ends));
						m_Interior = range.get_Interior();
						m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue))));
						starts = L"";
						ends = L"";
					}
				}
			}
		}
	}
	for (long r1 = m_Table1.FirstRowWithData; r1 <= m_Table1.NumberOfRows; r1++)
	{
		if (m_pbKeyMissing1[r1])
		{
			starts = convertR1C1(r1, 1);
			ends = convertR1C1(r1, m_Table1.NumberOfColumns);
			CRange range = m_Sheet1.get_Range(COleVariant(starts), COleVariant(ends));
			m_Interior = range.get_Interior();
			m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor1].red, m_Palette[m_nChosenColor1].green, m_Palette[m_nChosenColor1].blue))));
		}
	}
	for (long r2 = m_Table2.FirstRowWithData; r2 <= m_Table2.NumberOfRows; r2++)
	{
		if (m_pbKeyMissing2[r2]) // c1 - because we need it to run just once
		{
			starts = convertR1C1(r2, 1);
			ends = convertR1C1(r2, m_Table2.NumberOfColumns);
			CRange range = m_Sheet2.get_Range(COleVariant(starts), COleVariant(ends));
			m_Interior = range.get_Interior();
			m_Interior.put_Color(COleVariant(long(RGB(m_Palette[m_nChosenColor2].red, m_Palette[m_nChosenColor2].green, m_Palette[m_nChosenColor2].blue))));
		}
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 100);
	PostMessage(CM_UPDATE_PROGRESS2, 0, 100);
	m_bLockPrg2 = false;
	EndWaitCursor();
	if (g_pMainFrame) g_pMainFrame->updateStatusBar(CMsg(IDS_MARKING_DONE)); // CMsg(IDS_MARKING_DONE)
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
		column = m_nOldy;
		CString cnv = convertR1C1(row, column);
		CRange range = m_Sheet1.get_Range(COleVariant(cnv), COleVariant(cnv));
		m_Sheet1.Activate();
		range.Select();
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
		CRange range = m_Sheet2.get_Range(COleVariant(cnv), COleVariant(cnv));
		m_Sheet2.Activate();
		range.Select();
		if (m_bToFront)
		{
			m_App.put_Interactive(true);
			HWND ehWnd = (HWND)m_App.get_Hwnd();
			::PostMessage(ehWnd, WM_SHOWWINDOW, SW_RESTORE, 0);
			::SetForegroundWindow(ehWnd);
		}
	}
}
void CChildView::OnPut2front()
{
	m_bToFront = !m_bToFront;
}
void CChildView::OnUpdatePut2front(CCmdUI *pCmdUI)
{
	pCmdUI->SetCheck(m_bToFront);
}
void CChildView::suggestKeys1()
{
	int nPhase = 0;
	int attempts = 0;
	bool alreadyChecked = false;
	m_nCheckedKeysCounter1 = 0;
	for (int i = 0; i < SUGKEYS; i++)
	{
		m_nCheckedKeys1[i] = 0;
	}
	m_bLockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	m_nPossibleKeyCounter1 = 0;
	for (int i = 0; i < 255; i++)
	{
		m_nInvEntropy1[i] = 0;
	}
	for (int i = 0; i <= SUGKEYS + 1; i++)
	{
		m_nExaminedKeys1[i] = 0;
	}
	int	e_i;
	//////*********************
	long index[2];
	COleVariant vData;
	CString szdata;
	for (int i_h = 1; i_h <= m_Table1.NumberOfColumns; i_h++)
	{
		m_mapTmpMap1.clear();
		for (int i_i = m_Table1.FirstRowWithData; i_i <= m_Table1.NumberOfRows; i_i++)
		{
			szdata = "";
			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = i_h;
			try {
				m_saRet1.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata += vData;
			if ((m_mapTmpMap1.find(szdata) == m_mapTmpMap1.end()))
			{
				m_mapTmpMap1[szdata] = i_i;
				m_nInvEntropy1[i_h]++;
			}
		}
	}
	CalculateEntropyRank(1);
///////////********************
	if (m_Table1.NumberOfRows > 0)
	{
		int foundKeysSet1 = 10;
		while (true)
		{
			prgHlpr0 = attempts % 97; // only 3 keys
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
			}
			if (is2BExaminedOnce(1, SUGKEYS - 1))  // TODO: allow some zeros
			{
				alreadyChecked = getSimilarKeyProbability(1, SUGKEYS);
				if (!alreadyChecked)
				{
					foundKeysSet1 = createTempKeyArrays1();
					//CString tracestring = L"";
					//tracestring.Format(L"\n%i, %i, %i, %i, %i, %i, %i, %i, %i, %i : %i\n", examinedKeys1[0], examinedKeys1[1], examinedKeys1[2], examinedKeys1[3], examinedKeys1[4], examinedKeys1[5], examinedKeys1[6], examinedKeys1[7], examinedKeys1[8], examinedKeys1[9], foundKeysSet1);
					//TRACE(tracestring);
				}
			}
			else
			{
				foundKeysSet1 = 4; // Low entropy of key indexes
			}
			if (foundKeysSet1 == 0)
			{
				for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
				{
					m_PossibleKeys1[m_nPossibleKeyCounter1].k[tmp_i] = getNthEntropy(1, m_nExaminedKeys1[tmp_i]);
				}
				sortExaminedKeys(1);
				m_nPossibleKeyCounter1++;
			}
			if (attempts++ > m_nComplexity || m_nPossibleKeyCounter1 > m_Table1.NumberOfColumns /*  ( .....3 - sgn(table1.keys[0]) - sgn(table1.keys[1]) - sgn(table1.keys[2]))  */)
			{
				break; // quit in any case
			}
			e_i = SUGKEYS - 1;
			while (e_i >= 0)
			{
				/*if (table1.keys[e_i])
				{
					--e_i;
				}
				else */
				{
					if (m_nExaminedKeys1[e_i] >= m_Table1.NumberOfColumns)
					{
						m_nExaminedKeys1[e_i] = 0;
						--e_i; // decrement to the previous slot for key-column
					}
					else
					{
						++m_nExaminedKeys1[e_i];
						break;
					}
				}
			}
		}
	}
	else
	{
		MessageBox(CMsg(IDS_NO_SHEET_SELCTD_IN_FRST)); // TODO: CMsg(IDS_NO_SHEET_SELCTD_IN_FRST)
	}
	PostMessage(CM_UPDATE_PROGRESS, 0, 10000);
	return;
}
int CChildView::createTempKeyArrays1()
{
	long index[2];
	COleVariant vData;
	CString szdata;
	m_mapTmpMap1.clear();
	if (sumExaminedKeys(1, SUGKEYS - 1) > 0)
	{
		for (int i_i = m_Table1.FirstRowWithData; i_i <= m_Table1.NumberOfRows; i_i++)
		{
			szdata = "";
			for (int k_i = 0; k_i < SUGKEYS; k_i++)
			{
				if (m_nExaminedKeys1[k_i])
				{
					// Loop through the data and report the contents.
					index[0] = i_i;
					index[1] = getNthEntropy(1, m_nExaminedKeys1[k_i]);
					try {
						m_saRet1.GetElement(index, vData); vData = (CString)vData;
					}
					catch (COleException* e)
					{
						vData = L"";
					}
					szdata += vData;
				}
			}
			if (!(m_mapTmpMap1.find(szdata) == m_mapTmpMap1.end()))
			{
				m_mapTmpMap1.clear();
				return 2;
			}
			else
			{
				m_mapTmpMap1[szdata] = i_i;
			}
		}
	}
	else
	{
		return 1;
	}
	return 0;
	//search for the next available set of not-examined-yet columns
}
void CChildView::suggestKeys2()
{
	int nPhase = 0;
	int attempts = 0;
	bool alreadyChecked = false;
	m_nCheckedKeysCounter2 = 0;
	for (int i = 0; i < SUGKEYS; i++)
	{
		m_nCheckedKeys2[i] = 0;
	}
	m_bLockPrg2 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	m_nPossibleKeyCounter2 = 0;
	for (int i = 0; i < 255; i++)
	{
		m_nInvEntropy2[i] = 0;
	}
	for (int i = 0; i <= SUGKEYS + 1; i++)
	{
		m_nExaminedKeys2[i] = 0;
	} 
	int	e_i;
	//////*********************
	long index[2];
	COleVariant vData;
	CString szdata;
	for (int i_h = 1; i_h <= m_Table2.NumberOfColumns; i_h++)
	{
		m_mapTmpMap2.clear();
		for (int i_i = m_Table2.FirstRowWithData; i_i <= m_Table2.NumberOfRows; i_i++)
		{
			szdata = "";
			// Loop through the data and report the contents.
			index[0] = i_i;
			index[1] = i_h;
			try {
				m_saRet2.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata += vData;
			if ((m_mapTmpMap2.find(szdata) == m_mapTmpMap2.end()))
			{
				m_mapTmpMap2[szdata] = i_i;
				m_nInvEntropy2[i_h]++;
			}
		}
	}
	CalculateEntropyRank(2);
	///////////********************
	if (m_Table2.NumberOfRows > 0)
	{
		int foundKeysSet2 = 10;
		while (true)
		{
			prgHlpr0 = attempts % 97;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr);
				PostMessage(CM_UPDATE_KEYPROGRESS2, 0, prgHlpr);
			}
			if (is2BExaminedOnce(2, SUGKEYS - 1))  // TODO: allow some zeros
			{
				alreadyChecked = getSimilarKeyProbability(2, SUGKEYS);
				if (!alreadyChecked) foundKeysSet2 = createTempKeyArrays2();
			}
			else
			{
				foundKeysSet2 = 4; // Low entropy of key indexes
			}
			if (foundKeysSet2 == 0)
			{
				for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
				{
					m_PossibleKeys2[m_nPossibleKeyCounter2].k[tmp_i] = getNthEntropy(2, m_nExaminedKeys2[tmp_i]);
				}
				sortExaminedKeys(2);
				m_nPossibleKeyCounter2++;
			}
			if (attempts++ > m_nComplexity || m_nPossibleKeyCounter2 > m_Table2.NumberOfColumns /*( ... 3 - sgn(table2.keys[0]) - sgn(table2.keys[1]) - sgn(table2.keys[2])) */)
			{
				break; // quit in any case
			}
			e_i = SUGKEYS - 1;
			while (e_i >= 0)
			{
				/* if (table2.keys[e_i])
				{
					--e_i;
				}
				else */
				{
					if (m_nExaminedKeys2[e_i] >= m_Table2.NumberOfColumns)
					{
						m_nExaminedKeys2[e_i] = 0;
						--e_i; // decrement to the previous slot for key-column
					}
					else
					{
						++m_nExaminedKeys2[e_i];
						break;
					}
				}
			}
		}
	}
	else
	{
		MessageBox(CMsg(IDS_NO_SHEET_SELCTD_IN_SCND)); // TODO: CMsg(IDS_NO_SHEET_SELCTD_IN_SCND)
	}
	PostMessage(CM_UPDATE_PROGRESS2, 0, 20000);
	return;
}
int CChildView::createTempKeyArrays2()
{
	long index[2];
	COleVariant vData;
	CString szdata;
	m_mapTmpMap2.clear();
	if (sumExaminedKeys(2, SUGKEYS - 1) > 0)
	{
		for (int i_i = m_Table2.FirstRowWithData; i_i <= m_Table2.NumberOfRows; i_i++)
		{
			szdata = "";
			for (int k_i = 0; k_i < SUGKEYS; k_i++)
			{
				if (m_nExaminedKeys2[k_i])
				{
					// Loop through the data and report the contents.
					index[0] = i_i;
					index[1] = getNthEntropy(2, m_nExaminedKeys2[k_i]);;
					try {
						m_saRet2.GetElement(index, vData); vData = (CString)vData;
					}
					catch (COleException* e)
					{
						vData = L"";
					}
					szdata += vData;
				}
			}
			if (!(m_mapTmpMap2.find(szdata) == m_mapTmpMap2.end()))
			{
				m_mapTmpMap2.clear();
				return 2;
			}
			else
			{
				m_mapTmpMap2[szdata] = i_i;
			}
		}
	}
	else
	{
		return 1;
	}
	return 0;
	//search for the next available set of not-examined-yet columns
}
void CChildView::clearPossibleKeys()
{
	for (int i = 0; i < 255; i++)
	{
		for (int ii = 0; ii < SUGKEYS; ii++)
		{
			m_PossibleKeys1[i].k[ii] = 0;
			m_PossibleKeys2[i].k[ii] = 0;
		}
	}
	m_nPossibleKeyCounter1 = 0;
	m_nPossibleKeyCounter2 = 0;
}
inline void CChildView::sort3(int& a, int& b, int& c)
{
	if (c < b) swap(c, b);
	if (b < a) swap(b, a);
	if (c < b) swap(c, b);
}
UINT SuggestKeys1ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->suggestKeys1();
	AfxEndThread(0);
	return 0;
}
UINT SuggestKeys2ThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->suggestKeys2();
	AfxEndThread(0);
	return 0;
}
UINT MutualCheckThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->mutualCheck();
	AfxEndThread(0);
	return 0;
}
UINT FindSimsThreadProc(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->findSims();
	AfxEndThread(0);
	return 0;
}
UINT FindSimsThreadProc1(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->findSims1();
	AfxEndThread(0);
	return 0;
}
UINT FindSimsThreadProc2(LPVOID pParam)
{
	HWND hWnd1 = (HWND)pParam;
	CChildView* pWnd = (CChildView*)CWnd::FromHandle(hWnd1);
	pWnd->findSims2();
	AfxEndThread(0);
	return 0;
}
bool CChildView::mutualCheck()
{
	m_BestKeyComb.pk1 = 0;
	m_BestKeyComb.pk2 = 0;
	m_BestKeyComb.rating = 0;
	int tmpRslt = 0;
	// Cascade check
	if (m_pKeyProgressBar1 && m_pKeyProgressBar2)
	{
		PostMessage(CM_UPDATE_KEYPROGRESS1, 0, 0);
		PostMessage(CM_UPDATE_KEYPROGRESS2, 0, 0);
	}
	if (getNumberOfPossibleKeys(1, SUGKEYS, 0) == 0  && getNumberOfPossibleKeys(2, SUGKEYS, 0) == 0)
	{
		MessageBox(CMsg(IDS_NTHR_TBL_KEY_FND)); // CMsg(IDS_NTHR_TBL_KEY_FND)
		return false;
	}
	if (getNumberOfPossibleKeys(1, SUGKEYS, 0) == 0)
	{
		MessageBox(CMsg(IDS_NO_KEY_FND_IN_FRST)); // CMsg(IDS_NO_KEY_FND_IN_FRST)
		return false;
	}
	if (getNumberOfPossibleKeys(2, SUGKEYS, 0) == 0)
	{
		MessageBox(CMsg(IDS_NO_KEY_FND_IN_SCND)); // CMsg(IDS_NO_KEY_FND_IN_SCND)
		return false;
	}
	int m_i = 0;
	m_bLockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	int order = 1;
	while (m_i <= m_nPossibleKeyCounter1 && tmpRslt < 100)
	{
		prgHlpr0 = 100 * m_i / m_nPossibleKeyCounter1;
		if (prgHlpr0 > prgHlpr)
		{
			prgHlpr = prgHlpr0;
			PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
			PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
		}
		if (getNumberOfPossibleKeys(1, SUGKEYS, m_i) == order)
		{
			tmpRslt = checkKeys(m_i);
		}
		else
		{
			break;
		}
		m_i++;
	}
	order++;
	int maxOrder = getNumberOfPossibleKeys(1, SUGKEYS, (m_nPossibleKeyCounter1 - 1 >= 0 ? m_nPossibleKeyCounter1 - 1 : 0));
	while (tmpRslt < 90 && order <= maxOrder)
	{
		while (m_i <= m_nPossibleKeyCounter1 && tmpRslt < 90)
		{
			prgHlpr0 = 100 * m_i / m_nPossibleKeyCounter1;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS, 0, prgHlpr);
				PostMessage(CM_UPDATE_KEYPROGRESS1, 0, prgHlpr);
			}
			if (getNumberOfPossibleKeys(1, order, m_i) == order)
			{
				tmpRslt = checkKeys(m_i);
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
int CChildView::checkKeys(int tab1)
{
	long index[2];
	COleVariant vData;
	CString szdata;
	m_bLockPrg1 = true;
	int prgHlpr = 0, prgHlpr0 = 0;
	m_mapTmpMap1.clear();
	int order1 = getNumberOfPossibleKeys(1, SUGKEYS, tab1);
	if (order1 > 0)
	{
		for (int i_i = m_Table1.FirstRowWithData; i_i <= m_Table1.NumberOfRows; i_i++)
		{
			prgHlpr0 = 100 * i_i / m_Table1.NumberOfRows;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr );
				PostMessage(CM_UPDATE_KEYPROGRESS2, 0, prgHlpr);
			}
			szdata = "";
			for (int i_j = 0; i_j <= order1; i_j++)
			{
				if (m_PossibleKeys1[tab1].k[i_j])
				{
					// Loop through the data and report the contents.
					index[0] = i_i;
					index[1] = m_PossibleKeys1[tab1].k[i_j];
					try {
						m_saRet1.GetElement(index, vData); vData = (CString)vData;
					}
					catch (COleException* e)
					{
						vData = L"";
					}
					szdata += vData;
				}
			}
			m_mapTmpMap1[szdata] = i_i;
		}
	}
	else
	{ 
		return -1;  
	}
	int tab2 = 0;
	while (getNumberOfPossibleKeys(2, SUGKEYS, tab2) < order1)
	{
		tab2++;
	}
	long rslt = 0;
	while (getNumberOfPossibleKeys(2, SUGKEYS, tab2) == order1)
	{
		long found = 0;
		for (int i_i = m_Table2.FirstRowWithData; i_i <= m_Table2.NumberOfRows; i_i++)
		{
			prgHlpr0 = 100 * i_i / m_Table2.NumberOfRows;
			if (prgHlpr0 > prgHlpr)
			{
				prgHlpr = prgHlpr0;
				PostMessage(CM_UPDATE_PROGRESS2, 0, prgHlpr );
				PostMessage(CM_UPDATE_KEYPROGRESS2, 0, prgHlpr);
			}
			szdata = "";
			for (int i_k = 0; i_k <= order1; i_k++)
			{
				if (m_PossibleKeys2[tab2].k[i_k])
				{
					// Loop through the data and report the contents.
					index[0] = i_i;
					index[1] = m_PossibleKeys2[tab2].k[i_k];
					try {
						m_saRet2.GetElement(index, vData); vData = (CString)vData;
					}
					catch (COleException* e)
					{
						vData = L"";
					}
					szdata += vData;
				}
			}
			if (m_mapTmpMap1.count(szdata))
			{
				found++;
			}
		}
		rslt = 100 * found / min(m_Table2.NumberOfRows - m_Table2.FirstRowWithData, m_Table1.NumberOfRows - m_Table1.FirstRowWithData);
		if (rslt > m_BestKeyComb.rating)
		{
			m_BestKeyComb.pk1 = tab1;
			m_BestKeyComb.pk2 = tab2;
			m_BestKeyComb.rating = rslt;
			m_BestKeyComb.cnt = found;
		}
		tab2++;
	}
	if (m_pKeyProgressBar1 && m_pKeyProgressBar2)
	{
		PostMessage(CM_UPDATE_KEYPROGRESS1, 0, 0);
		PostMessage(CM_UPDATE_KEYPROGRESS2, 0, 0);
	}
	return m_BestKeyComb.rating;
}
int CChildView::deleteKey(int table, int column)
{
	int rslt = 0;
	int tmp_i = 0;
	if (table == 1)
	{
		for (tmp_i = 0; tmp_i < m_nKeyPairCounter; tmp_i++)
		{
			if (m_KeyPair[tmp_i].tab1 == column)
			{
				deleteKeyAt(tmp_i);
				rslt++;
			}
		}
	}
	if (table == 2)
	{
		for (tmp_i = 0; tmp_i < m_nKeyPairCounter; tmp_i++)
		{
			if (m_KeyPair[tmp_i].tab2 == column)
			{
				deleteKeyAt(tmp_i);
				rslt++;
			}
		}
	}
	return rslt;
}
void CChildView::setKey(int table, int column)
{
	if (table == 1)
	{
		m_KeyPair[m_nKeyPairCounter].tab1 = column;
	}
	if (table == 2)
	{
		m_KeyPair[m_nKeyPairCounter].tab2 = column;
	}
}
void CChildView::deleteAllKeys()
{
	for (int i_i = 1; i_i < m_nKeyPairCounter; i_i++)
	{
		m_KeyPair[i_i].tab1 = 0;
		m_KeyPair[i_i].tab2 = 0;
	}
	m_nKeyPairCounter = 0;
	m_bToDisplaySimilarClms = false;
	m_bXSimilarityComputed = false;
	m_vecSimilaritiesAcrossTables.clear();
	m_vecSimilaritiesAcrossTablesSorted.clear();
}
bool CChildView::areThereAnyKeys()
{
	return m_nKeyPairCounter == 0 ? false : true;
}
bool CChildView::isThisAKey(int table, int column)
{
	if (table == 1)
	{
		for (int tmp_i = 0; tmp_i < m_nKeyPairCounter; tmp_i++)
		{
			if (m_KeyPair[tmp_i].tab1 == column)
			{
				return true;
			}
		}
		return false;
	}
	if (table == 2)
	{
		for (int tmp_i = 0; tmp_i < m_nKeyPairCounter; tmp_i++)
		{
			if (m_KeyPair[tmp_i].tab2 == column)
			{
				return true;
			}
		}
		return false;
	}
}
int CChildView::getNthKey(int table, int key)
{
	if (table == 1)
	{
		return m_KeyPair[key].tab1;
	}
	else
	{
		return m_KeyPair[key].tab2;
	}
}
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
void CChildView::setNthKey(int n, int col1, int col2)
{
	m_KeyPair[n].tab1 = col1;
	m_KeyPair[n].tab2 = col2;
}
void CChildView::insertKeyAt(int n, int col1, int col2)
{
	for (int tmp_i = m_nKeyPairCounter; tmp_i > n; tmp_i--)
	{
		m_KeyPair[tmp_i].tab1 = m_KeyPair[tmp_i - 1].tab1;
		m_KeyPair[tmp_i].tab2 = m_KeyPair[tmp_i - 1].tab2;
	}
	m_nKeyPairCounter++;
}
void CChildView::deleteKeyAt(int n)
{
	for (int tmp_i = n; tmp_i < m_nKeyPairCounter; tmp_i++)
	{
		m_KeyPair[tmp_i].tab1 = m_KeyPair[tmp_i + 1].tab1;
		m_KeyPair[tmp_i].tab2 = m_KeyPair[tmp_i + 1].tab2;
	}
	m_nKeyPairCounter--;
}
void CChildView::pushKey(int col1, int col2)
{
	m_KeyPair[m_nKeyPairCounter].tab1 = col1;
	m_KeyPair[m_nKeyPairCounter].tab2 = col2;
	m_nKeyPairCounter++;
}
bool CChildView::usePossibleKeys()
{
	deleteAllKeys();
	int numberOfPossibleKeys = getNumberOfPossibleKeys();
	for (int tmp_i = 0; tmp_i <= numberOfPossibleKeys; tmp_i++)
	{
		if (m_PossibleKeys1[m_BestKeyComb.pk1].k[tmp_i] + m_PossibleKeys2[m_BestKeyComb.pk2].k[tmp_i])
		{
			pushKey(m_PossibleKeys1[m_BestKeyComb.pk1].k[tmp_i], m_PossibleKeys2[m_BestKeyComb.pk2].k[tmp_i]);
		}
	}
	return false;
}
int CChildView::getNumberOfPossibleKeys()
{
	for (int tmp_i = 1; tmp_i < 255; tmp_i++)
	{
		if (m_PossibleKeys1[m_BestKeyComb.pk1].k[tmp_i] == 0 && m_PossibleKeys2[m_BestKeyComb.pk2].k[tmp_i] == 0)
		{
			return tmp_i;
		}
	}
	return 0;
}
void CChildView::sortExaminedKeys(int table)
{
	int nonzerosNr = 0;
	if (table == 1)
	{
		for (int i_i = 0; i_i < SUGKEYS; i_i++)
		{
			if (m_PossibleKeys1[m_nPossibleKeyCounter1].k[i_i] > 0 && nonzerosNr < i_i)
			{
				m_PossibleKeys1[m_nPossibleKeyCounter1].k[nonzerosNr] = m_PossibleKeys1[m_nPossibleKeyCounter1].k[i_i];
				m_PossibleKeys1[m_nPossibleKeyCounter1].k[i_i] = 0;
				nonzerosNr++;
			}
		}
	}
	else
	{
		for (int i_i = 0; i_i < SUGKEYS; i_i++)
		{
			if (m_PossibleKeys2[m_nPossibleKeyCounter2].k[i_i] > 0 && nonzerosNr < i_i)
			{
				m_PossibleKeys2[m_nPossibleKeyCounter2].k[nonzerosNr] = m_PossibleKeys2[m_nPossibleKeyCounter2].k[i_i];
				m_PossibleKeys2[m_nPossibleKeyCounter2].k[i_i] = 0;
				nonzerosNr++;
			}
		}
	}
// return nonzerosNr;
}
int CChildView::sumExaminedKeys(int table, int nmax)
{
	int rslt = 0;
	if (table == 1)
	{
		for (int tmp_i = 0; tmp_i <= nmax; tmp_i++)
		{
			rslt += m_nExaminedKeys1[tmp_i];
		}
	}
	else
	{
		for (int tmp_i = 0; tmp_i <= nmax; tmp_i++)
		{
			rslt += m_nExaminedKeys2[tmp_i];
		}
	}
	return rslt;
}
bool CChildView::is2BExaminedOnce(int table, int max)
{
	if (table == 1)
	{
		for (int tmp_i0 = 0; tmp_i0 <= max; tmp_i0++)
		{
			if (m_nExaminedKeys1[tmp_i0] > 0)
			{
				for (int tmp_i1 = tmp_i0 + 1; tmp_i1 <= max; tmp_i1++)
				{
					if (m_nExaminedKeys1[tmp_i0] == m_nExaminedKeys1[tmp_i1])
					{
						return false;
					}
				}
			}
		}
	}
	else
	{
		for (int tmp_i0 = 0; tmp_i0 <= max; tmp_i0++)
		{
			if (m_nExaminedKeys2[tmp_i0] > 0)
			{
				for (int tmp_i1 = tmp_i0 + 1; tmp_i1 <= max; tmp_i1++)
				{
					if (m_nExaminedKeys2[tmp_i0] == m_nExaminedKeys2[tmp_i1])
					{
						return false;
					}
				}
			}
		}
	}
	return true;
}
bool CChildView::getSimilarKeyProbability(int table, int max)
{
	int similarKeyProbab = 0;
	int tmpa[SUGKEYS];
	unsigned long long ullTest = 0;
	if (table == 1)
	{
		for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
		{
			if (m_nExaminedKeys1[tmp_i])
			{
				ullTest += pow(2, m_nExaminedKeys1[tmp_i]);
			}
		}
		for (int tmp_i0 = 0; tmp_i0 <= m_nCheckedKeysCounter1; tmp_i0++)
		{
			if (m_nCheckedKeys1[tmp_i0] == ullTest)
			{
				return true;
			}
		}
		if (m_nCheckedKeysCounter1 < m_nComplexity)
		{
			m_nCheckedKeys1[m_nCheckedKeysCounter1] = ullTest;
			m_nCheckedKeysCounter1++;
		}
		else
		{
			return true;
		}
	}
	else
	{
		for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
		{
			if (m_nExaminedKeys2[tmp_i])
			{
				ullTest += pow(2, m_nExaminedKeys2[tmp_i]);
			}
		}
		for (int tmp_i0 = 0; tmp_i0 <= m_nCheckedKeysCounter2; tmp_i0++)
		{
			if (m_nCheckedKeys2[tmp_i0] == ullTest)
			{
				return true;
			}
		}
		if (m_nCheckedKeysCounter2 < m_nComplexity)
		{
			m_nCheckedKeys2[m_nCheckedKeysCounter2] = ullTest;
			m_nCheckedKeysCounter2++;
		}
		else
		{
			return true;
		}
	}
	return false;
}
int CChildView::getNthEntropy(int table, int n)
{
	if (table == 1)
	{
		return m_nSortedEntropy1[n];
	}
	else
	{
		return m_nSortedEntropy2[n];
	}
	return 0;
}
int CChildView::CalculateEntropyRank(int table)
{
	int hlpr_index;
	int hlpr_value;
	int lasthlpr_value = 0;
	int stored = 0;
	int i2 = 0;
	i2 = 1;
	if (table == 1)
	{
		for (int i0 = 0; i0 < 255; i0++)
		{
			m_nSortedEntropy1[i0] = 0;
		}
		while (stored < m_Table1.NumberOfColumns)
		{
			hlpr_index = 0;
			hlpr_value = 0;
			for (int i1 = 1; i1 <= m_Table1.NumberOfColumns; i1++)
			{
				if (m_nInvEntropy1[i1] >= hlpr_value && !isEntropyStored(1, i1, stored))
				{
					hlpr_value = m_nInvEntropy1[i1];
					hlpr_index = i1;
				}
			}
			if (hlpr_index > 0)
			{
				stored++;
				m_nSortedEntropy1[stored] = hlpr_index;
			}
		}
	}
	else
	{
		for (int i0 = 0; i0 < 255; i0++)
		{
			m_nSortedEntropy2[i0] = 0;
		}
		while (stored < m_Table2.NumberOfColumns)
		{
			hlpr_index = 0;
			hlpr_value = 0;
			for (int i2 = 1; i2 <= m_Table2.NumberOfColumns; i2++)
			{
				if (m_nInvEntropy2[i2] >= hlpr_value && !isEntropyStored(2, i2, stored))
				{
					hlpr_value = m_nInvEntropy2[i2];
					hlpr_index = i2;
				}
			}
			if (hlpr_index > 0)
			{
				stored++;
				m_nSortedEntropy2[stored] = hlpr_index;
			}
		}
	}
	return 0;
}
bool CChildView::isEntropyStored(int table, int clm, int max)
{
	if (table == 1)
	{
		for (int i = 1; i <= max; i++)
		{
			if (m_nSortedEntropy1[i] == clm)
			{
				return true;
			}
		}
	}
	else
	{
		for (int i = 1; i <= max; i++)
		{
			if (m_nSortedEntropy2[i] == clm)
			{
				return true;
			}
		}
	}
	return false;
}
void CChildView::OnUpdateCombo2(CCmdUI *pCmdUI)
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
void CChildView::OnCombo2()
{
	if (m_pCombo2->GetCurSel() == 0) m_nComplexity = 10000;
	if (m_pCombo2->GetCurSel() == 1) m_nComplexity = 100000;
	if (m_pCombo2->GetCurSel() == 2) m_nComplexity = 1000000;
}
//int CChildView::getPossibleKeyReadiness(int table, int order)
//{
//
//	return 0;
//}
int CChildView::getNumberOfPossibleKeys(int table, int order, int item)
{
	int cnt = 0;
	if (table == 1)
	{
		for (int i = 0; i < order; i++)
		{
			cnt += sgn(m_PossibleKeys1[item].k[i]);
		}
	}
	else
	{
		for (int i = 0; i < order; i++)
		{
			cnt += sgn(m_PossibleKeys2[item].k[i]);
		}
	}
	return cnt;
}
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
			index[0] = r_i1;
			index[1] = c_i1;
			try {
				m_saRet1.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata = vData;
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
				index[0] = r_i2;
				index[1] = c_i2;
				try {
					m_saRet2.GetElement(index, vData); vData = (CString)vData;
				}
				catch (COleException* e)
				{
					vData = L"";
				}
				szdata = vData;
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
			index[0] = r_i1;
			index[1] = c_i1;
			try {
				m_saRet1.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata = vData;
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
				index[0] = r_i2;
				index[1] = c_i2;
				try {
					m_saRet2.GetElement(index, vData); vData = (CString)vData;
				}
				catch (COleException* e)
				{
					vData = L"";
				}
				szdata = vData;
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
			index[0] = r_i1;
			index[1] = c_i1;
			try {
				m_saTmpRet1.GetElement(index, vData); vData = (CString)vData;
			}
			catch (COleException* e)
			{
				vData = L"";
			}
			szdata = vData;
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
				index[0] = r_i2;
				index[1] = c_i2;
				try {
					m_saTmpRet2.GetElement(index, vData); vData = (CString)vData;
				}
				catch (COleException* e)
				{
					vData = L"";
				}
				szdata = vData;
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
void CChildView::OnUpdateSimilarpaircheckbox(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(m_Table1.NumberOfColumns * m_Table2.NumberOfColumns && m_bXSimilarityComputed);
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pShowSims = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_SIMILARPAIRCHECKBOX));
	pCmdUI->SetCheck(m_bToDisplaySimilarClms);
}
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
	HWND hWnd0 = this->GetSafeHwnd();
	//AfxBeginThread(FindSimsThreadProc, hWnd0);
	AfxBeginThread(FindSimsThreadProc1, hWnd0);
	AfxBeginThread(FindSimsThreadProc2, hWnd0);
	m_bLockPrg1 = true;
	m_bLockPrg2 = true;
}
void CChildView::OnIdxcrtBtn()
{
}
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
void CChildView::OnUpdateKeyProgress1(CCmdUI *pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pKeyProgressBar1 = DYNAMIC_DOWNCAST(CMFCRibbonProgressBar, m_pRibbon->FindByID(ID_KEY_PROGRESS1));
}
void CChildView::OnUpdateKeyProgress2(CCmdUI *pCmdUI)
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
void CChildView::OnUpdateIdxCheckbox(CCmdUI *pCmdUI)
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
		if (startpos == -1 || startpos >= len) startpos = len - 1;
		for (LPCTSTR lpszReverse = lpszData + startpos;
			lpszReverse != lpszData; --lpszReverse)
			if (_tcsncmp(lpszSub, lpszReverse, lenSub) == 0)
				return (lpszReverse - lpszData);
	}
	return -1;
}
void CChildView::OnCheckIdx()
{
	m_bUseIndexes = !m_bUseIndexes;
}
void CChildView::OnUpdateCheckIdx(CCmdUI *pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_IDX_CHECKBOX));
	pCmdUI->SetCheck(m_bUseIndexes);
}
void CChildView::OnUsidxCheck()
{
	m_bUseIndexes = !m_bUseIndexes;
	if (m_bUseIndexes) MessageBox(CMsg(IDS_IDXING_WARNING)); // CMsg(IDS_IDXING_WARNING)
}
void CChildView::OnUpdateUsidxCheck(CCmdUI *pCmdUI)
{
	m_pRibbon = ((CFrameWndEx*)AfxGetMainWnd())->GetRibbonBar();
	m_pUseIndices = DYNAMIC_DOWNCAST(CMFCRibbonCheckBox, m_pRibbon->FindByID(ID_USIDX_CHECK));
	pCmdUI->SetCheck(m_bUseIndexes);
}
void CChildView::OnUpdateRows1(CCmdUI *pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
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
void CChildView::OnUpdateCols1(CCmdUI *pCmdUI)
{
	if (!(m_szFilename1 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
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
void CChildView::OnUpdateRows2(CCmdUI *pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
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
void CChildView::OnUpdateCols2(CCmdUI *pCmdUI)
{
	if (!(m_szFilename2 == "")) pCmdUI->Enable(true); else pCmdUI->Enable(false);
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